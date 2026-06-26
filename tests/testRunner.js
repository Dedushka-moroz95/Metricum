(function (global) {
  const tests = [];

  function test(name, run) {
    tests.push({ name: name, run: run });
  }

  function assert(condition, message) {
    if (!condition) {
      throw new Error(message || "Assertion failed");
    }
  }

  assert.equal = function (actual, expected, message) {
    if (actual !== expected) {
      throw new Error((message || "Values are not equal") + "\nExpected: " + expected + "\nActual: " + actual);
    }
  };

  assert.closeTo = function (actual, expected, precision, message) {
    const safePrecision = precision === undefined ? 0.000001 : precision;

    if (Math.abs(actual - expected) > safePrecision) {
      throw new Error((message || "Values are not close") + "\nExpected: " + expected + "\nActual: " + actual);
    }
  };

  assert.deepEqual = function (actual, expected, message) {
    const actualText = JSON.stringify(actual);
    const expectedText = JSON.stringify(expected);

    if (actualText !== expectedText) {
      throw new Error((message || "Objects are not equal") + "\nExpected: " + expectedText + "\nActual: " + actualText);
    }
  };

  function createRow(rowNumber, values) {
    return {
      rowNumber: rowNumber,
      values: values,
    };
  }

  function createTable(rows, headers) {
    return {
      fileName: "report.xlsx",
      sheetName: "Sheet1",
      headerRowIndex: 1,
      headers:
        headers || [
          { id: "period", name: "Period" },
          { id: "name", name: "Name" },
          { id: "sales", name: "Sales" },
          { id: "quality", name: "Quality %" },
        ],
      rows: rows,
      previewRows: rows.slice(0, 8),
      warnings: [],
    };
  }

  function createPeriod(id, label, rows, idColumn) {
    return {
      id: id,
      label: label,
      file: { name: label + ".xlsx" },
      idColumn: idColumn || "name",
      table: createTable(rows),
    };
  }

  function createMetric(id, label, columnId, aggregation) {
    return {
      id: id,
      label: label,
      aggregation: aggregation || "auto",
      columns: {
        p1: columnId,
        p2: columnId,
        p3: columnId,
      },
    };
  }

  function createMetricWithColumns(id, label, columns, aggregation) {
    return {
      id: id,
      label: label,
      aggregation: aggregation || "auto",
      columns: columns,
    };
  }

  function findRow(comparison, label) {
    return comparison.rows.find(function (row) {
      return row.label === label;
    });
  }

  function findMetricResult(row, metricId) {
    return row.metrics.find(function (metric) {
      return metric.metricId === metricId;
    });
  }

  function registerTests() {
    const App = global.Metricum || {};
    const Normalizers = App.Normalizers;
    const PeriodBuilder = App.PeriodBuilder;
    const Comparator = App.Comparator;

    test("Модули ядра загружены", function () {
      assert(Normalizers, "Normalizers не найден");
      assert(PeriodBuilder, "PeriodBuilder не найден");
      assert(Comparator, "Comparator не найден");
    });

    test("normalizeKey приводит текст к стабильному ключу", function () {
      assert.equal(Normalizers.normalizeKey("  Anna   Kuznetsova  "), "anna kuznetsova");
    });

    test("normalizeNumber читает числа с пробелами, запятыми и процентом", function () {
      const result = Normalizers.normalizeNumber("1 234,56%");

      assert.equal(result.isNumeric, true);
      assert.closeTo(result.value, 1234.56);
    });

    test("normalizeNumber читает европейский формат с точкой и запятой", function () {
      const result = Normalizers.normalizeNumber("1.234,56");

      assert.equal(result.isNumeric, true);
      assert.closeTo(result.value, 1234.56);
    });

    test("normalizeNumber читает английский формат с запятой и точкой", function () {
      const result = Normalizers.normalizeNumber("1,234.56");

      assert.equal(result.isNumeric, true);
      assert.closeTo(result.value, 1234.56);
    });

    test("normalizeNumber читает отрицательные значения в скобках", function () {
      const result = Normalizers.normalizeNumber("(42)");

      assert.equal(result.isNumeric, true);
      assert.equal(result.value, -42);
    });

    test("normalizeNumber отличает пустые значения от нечислового текста", function () {
      const empty = Normalizers.normalizeNumber("-");
      const text = Normalizers.normalizeNumber("ошибка");

      assert.equal(empty.isEmpty, true);
      assert.equal(empty.isNumeric, false);
      assert.equal(text.isEmpty, false);
      assert.equal(text.isNumeric, false);
    });

    test("isEmptyValue распознает основные пустые Excel-значения", function () {
      assert.equal(Normalizers.isEmptyValue(""), true);
      assert.equal(Normalizers.isEmptyValue("   "), true);
      assert.equal(Normalizers.isEmptyValue("-"), true);
      assert.equal(Normalizers.isEmptyValue("n/a"), true);
      assert.equal(Normalizers.isEmptyValue("нет"), true);
    });

    test("percent helpers распознают процентные заголовки и значения", function () {
      assert.equal(Normalizers.headerLooksPercent("Quality %"), true);
      assert.equal(Normalizers.headerLooksPercent("CSAT rate"), true);
      assert.equal(Normalizers.valueLooksPercent("92%"), true);
      assert.equal(Normalizers.valueLooksPercent("92"), false);
    });

    test("calculatePercentChange считает изменение относительно прошлого периода", function () {
      assert.equal(Normalizers.calculatePercentChange(100, 125), 25);
      assert.equal(Normalizers.calculatePercentChange(0, 125), null);
    });

    test("PeriodBuilder собирает виртуальные периоды из строк одного файла", function () {
      const table = createTable([
        createRow(2, { period: "January", name: "Anna", sales: "10" }),
        createRow(3, { period: "February", name: "Anna", sales: "15" }),
        createRow(4, { period: "January", name: "Boris", sales: "20" }),
      ]);
      const result = PeriodBuilder.buildVirtualPeriods({
        table: table,
        periodColumn: "period",
      });

      assert.equal(result.periods.length, 2);
      assert.deepEqual(
        result.periods.map(function (period) { return period.label; }),
        ["January", "February"]
      );
      assert.equal(result.periods[0].source.rowCount, 2);
      assert.equal(result.periods[1].source.rowCount, 1);
    });

    test("PeriodBuilder объединяет одинаковые периоды после trim и lowercase", function () {
      const table = createTable([
        createRow(2, { period: " January ", name: "Anna", sales: "10" }),
        createRow(3, { period: "january", name: "Boris", sales: "20" }),
        createRow(4, { period: "February", name: "Anna", sales: "15" }),
      ]);
      const result = PeriodBuilder.buildVirtualPeriods({
        table: table,
        periodColumn: "period",
      });

      assert.equal(result.periods.length, 2);
      assert.equal(result.periods[0].label, "January");
      assert.equal(result.periods[0].table.rows.length, 2);
      assert.equal(result.periods[0].source.periodValue, "January");
    });

    test("PeriodBuilder сохраняет метаданные виртуального периода", function () {
      const table = createTable([
        createRow(2, { period: "January", name: "Anna", sales: "10" }),
        createRow(3, { period: "January", name: "Boris", sales: "20" }),
        createRow(4, { period: "February", name: "Anna", sales: "15" }),
      ]);
      const result = PeriodBuilder.buildVirtualPeriods({
        table: table,
        periodColumn: "period",
      });
      const period = result.periods[0];

      assert.equal(period.source.type, "singleFile");
      assert.equal(period.source.fileName, "report.xlsx");
      assert.equal(period.source.periodColumn, "period");
      assert.equal(period.table.previewRows.length, 2);
    });

    test("PeriodBuilder предупреждает о пустых значениях в колонке периода", function () {
      const table = createTable([
        createRow(2, { period: "January", name: "Anna", sales: "10" }),
        createRow(3, { period: "", name: "Boris", sales: "20" }),
        createRow(4, { period: "February", name: "Anna", sales: "15" }),
      ]);
      const result = PeriodBuilder.buildVirtualPeriods({
        table: table,
        periodColumn: "period",
      });

      assert.equal(result.periods.length, 2);
      assert(result.warnings.length > 0, "Ожидалось предупреждение о пустом периоде");
    });

    test("PeriodBuilder предупреждает, если периодов меньше двух", function () {
      const table = createTable([
        createRow(2, { period: "January", name: "Anna", sales: "10" }),
      ]);
      const result = PeriodBuilder.buildVirtualPeriods({
        table: table,
        periodColumn: "period",
      });

      assert.equal(result.periods.length, 1);
      assert(result.warnings.length > 0, "Ожидалось предупреждение");
    });

    test("PeriodBuilder возвращает пустой результат без таблицы или колонки периода", function () {
      const result = PeriodBuilder.buildVirtualPeriods({
        table: null,
        periodColumn: "",
      });

      assert.equal(result.periods.length, 0);
      assert.equal(result.warnings.length, 0);
    });

    test("Comparator считает endpoint-динамику между первым и последним периодом", function () {
      const periods = [
        createPeriod("p1", "January", [
          createRow(2, { name: "Anna", sales: "10" }),
          createRow(3, { name: "Boris", sales: "20" }),
        ]),
        createPeriod("p2", "February", [
          createRow(2, { name: "Anna", sales: "15" }),
          createRow(3, { name: "Boris", sales: "18" }),
        ]),
      ];
      const comparison = Comparator.comparePeriods({
        periods: periods,
        metrics: [createMetric("sales", "Sales", "sales")],
        comparisonMode: "endpoint",
      });
      const result = findMetricResult(findRow(comparison, "Anna"), "sales");

      assert.equal(comparison.rows.length, 2);
      assert.equal(result.valueA, 10);
      assert.equal(result.valueB, 15);
      assert.equal(result.delta, 5);
      assert.equal(result.deltaPercent, 50);
      assert.equal(result.impact, "good");
    });

    test("Comparator считает последовательные сравнения между соседними периодами", function () {
      const periods = [
        createPeriod("p1", "January", [createRow(2, { name: "Anna", sales: "10" })]),
        createPeriod("p2", "February", [createRow(2, { name: "Anna", sales: "15" })]),
        createPeriod("p3", "March", [createRow(2, { name: "Anna", sales: "12" })]),
      ];
      const comparison = Comparator.comparePeriods({
        periods: periods,
        metrics: [createMetric("sales", "Sales", "sales")],
        comparisonMode: "sequential",
      });
      const comparisons = findMetricResult(findRow(comparison, "Anna"), "sales").comparisons;

      assert.equal(comparisons.length, 2);
      assert.equal(comparisons[0].label, "February - January");
      assert.equal(comparisons[0].delta, 5);
      assert.equal(comparisons[0].impact, "good");
      assert.equal(comparisons[1].label, "March - February");
      assert.equal(comparisons[1].delta, -3);
      assert.equal(comparisons[1].impact, "bad");
    });

    test("Comparator фиксирует отсутствующие объекты по периодам", function () {
      const periods = [
        createPeriod("p1", "January", [
          createRow(2, { name: "Anna", sales: "10" }),
          createRow(3, { name: "Boris", sales: "20" }),
        ]),
        createPeriod("p2", "February", [
          createRow(2, { name: "Anna", sales: "15" }),
        ]),
      ];
      const comparison = Comparator.comparePeriods({
        periods: periods,
        metrics: [createMetric("sales", "Sales", "sales")],
        comparisonMode: "endpoint",
      });

      assert.equal(comparison.missingByPeriod[1].items.length, 1);
      assert.equal(comparison.missingByPeriod[1].items[0].label, "Boris");
      assert.equal(findRow(comparison, "Boris").isComplete, false);
    });

    test("Comparator фиксирует пустые идентификаторы", function () {
      const periods = [
        createPeriod("p1", "January", [
          createRow(2, { name: "", sales: "10" }),
          createRow(3, { name: "Anna", sales: "20" }),
        ]),
        createPeriod("p2", "February", [
          createRow(2, { name: "Anna", sales: "25" }),
        ]),
      ];
      const comparison = Comparator.comparePeriods({
        periods: periods,
        metrics: [createMetric("sales", "Sales", "sales")],
        comparisonMode: "endpoint",
      });

      assert.equal(comparison.emptyIdsByPeriod[0].items.length, 1);
      assert.equal(comparison.emptyIdsByPeriod[0].items[0].rowNumber, 2);
    });

    test("Comparator агрегирует дубли по сумме для обычных чисел", function () {
      const periods = [
        createPeriod("p1", "January", [
          createRow(2, { name: "Anna", sales: "10" }),
          createRow(3, { name: "Anna", sales: "5" }),
        ]),
        createPeriod("p2", "February", [
          createRow(2, { name: "Anna", sales: "20" }),
        ]),
      ];
      const comparison = Comparator.comparePeriods({
        periods: periods,
        metrics: [createMetric("sales", "Sales", "sales")],
        comparisonMode: "endpoint",
      });
      const result = comparison.rows[0].metrics[0];

      assert.equal(result.periodValues[0].value, 15);
      assert.equal(result.delta, 5);
      assert.equal(comparison.duplicatesByPeriod[0].items[0].count, 2);
    });

    test("Comparator поддерживает явную агрегацию дублей через average", function () {
      const periods = [
        createPeriod("p1", "January", [
          createRow(2, { name: "Anna", sales: "10" }),
          createRow(3, { name: "Anna", sales: "20" }),
        ]),
        createPeriod("p2", "February", [
          createRow(2, { name: "Anna", sales: "25" }),
        ]),
      ];
      const comparison = Comparator.comparePeriods({
        periods: periods,
        metrics: [createMetric("sales", "Sales", "sales", "avg")],
        comparisonMode: "endpoint",
      });
      const result = comparison.rows[0].metrics[0];

      assert.equal(result.periodValues[0].aggregation, "avg");
      assert.equal(result.periodValues[0].value, 15);
      assert.equal(result.delta, 10);
    });

    test("Comparator поддерживает агрегацию min, max и first", function () {
      const periods = [
        createPeriod("p1", "January", [
          createRow(2, { name: "Anna", sales: "10" }),
          createRow(3, { name: "Anna", sales: "20" }),
          createRow(4, { name: "Anna", sales: "5" }),
        ]),
        createPeriod("p2", "February", [
          createRow(2, { name: "Anna", sales: "30" }),
        ]),
      ];
      const metrics = [
        createMetric("min_sales", "Min Sales", "sales", "min"),
        createMetric("max_sales", "Max Sales", "sales", "max"),
        createMetric("first_sales", "First Sales", "sales", "first"),
      ];
      const comparison = Comparator.comparePeriods({
        periods: periods,
        metrics: metrics,
        comparisonMode: "endpoint",
      });
      const row = comparison.rows[0];

      assert.equal(findMetricResult(row, "min_sales").periodValues[0].value, 5);
      assert.equal(findMetricResult(row, "max_sales").periodValues[0].value, 20);
      assert.equal(findMetricResult(row, "first_sales").periodValues[0].value, 10);
    });

    test("Comparator усредняет процентные дубли в режиме auto", function () {
      const periods = [
        createPeriod("p1", "January", [
          createRow(2, { name: "Anna", quality: "90%" }),
          createRow(3, { name: "Anna", quality: "80%" }),
        ]),
        createPeriod("p2", "February", [
          createRow(2, { name: "Anna", quality: "100%" }),
        ]),
      ];
      const comparison = Comparator.comparePeriods({
        periods: periods,
        metrics: [createMetric("quality", "Quality", "quality")],
        comparisonMode: "endpoint",
      });
      const result = comparison.rows[0].metrics[0];

      assert.equal(result.valueFormat, "percent");
      assert.equal(result.periodValues[0].value, 85);
      assert.equal(result.delta, 15);
    });

    test("Comparator масштабирует доли до процентов по процентному заголовку", function () {
      const periods = [
        createPeriod("p1", "January", [
          createRow(2, { name: "Anna", quality: 0.9 }),
        ]),
        createPeriod("p2", "February", [
          createRow(2, { name: "Anna", quality: 0.95 }),
        ]),
      ];
      const comparison = Comparator.comparePeriods({
        periods: periods,
        metrics: [createMetric("quality", "Quality", "quality")],
        comparisonMode: "endpoint",
      });
      const result = comparison.rows[0].metrics[0];

      assert.equal(result.valueFormat, "percent");
      assert.equal(result.valueA, 90);
      assert.equal(result.valueB, 95);
      assert.closeTo(result.delta, 5);
    });

    test("Comparator фиксирует нечисловые значения показателей", function () {
      const periods = [
        createPeriod("p1", "January", [
          createRow(2, { name: "Anna", sales: "abc" }),
        ]),
        createPeriod("p2", "February", [
          createRow(2, { name: "Anna", sales: "20" }),
        ]),
      ];
      const comparison = Comparator.comparePeriods({
        periods: periods,
        metrics: [createMetric("sales", "Sales", "sales")],
        comparisonMode: "endpoint",
      });
      const result = comparison.rows[0].metrics[0];

      assert.equal(comparison.invalidValues.length, 1);
      assert.equal(result.valueA, null);
      assert.equal(result.delta, null);
      assert.equal(result.impact, "unknown");
    });

    test("Comparator считает ручные пары периодов", function () {
      const periods = [
        createPeriod("p1", "January", [createRow(2, { name: "Anna", sales: "10" })]),
        createPeriod("p2", "February", [createRow(2, { name: "Anna", sales: "15" })]),
        createPeriod("p3", "March", [createRow(2, { name: "Anna", sales: "21" })]),
      ];
      const comparison = Comparator.comparePeriods({
        periods: periods,
        metrics: [createMetric("sales", "Sales", "sales")],
        comparisonMode: "manual",
        comparisonPairs: [
          {
            fromPeriodId: "p1",
            toPeriodId: "p2",
            label: "February - January",
          },
          {
            fromPeriodId: "p2",
            toPeriodId: "p3",
            label: "March - February",
          },
        ],
      });
      const comparisons = comparison.rows[0].metrics[0].comparisons;

      assert.equal(comparisons.length, 2);
      assert.equal(comparisons[0].delta, 5);
      assert.equal(comparisons[1].delta, 6);
    });

    test("Comparator поддерживает разные колонки показателя по периодам", function () {
      const periods = [
        createPeriod("p1", "January", [
          createRow(2, { name: "Anna", sales_old: "10" }),
        ]),
        createPeriod("p2", "February", [
          createRow(2, { name: "Anna", sales_new: "14" }),
        ]),
      ];
      const comparison = Comparator.comparePeriods({
        periods: periods,
        metrics: [
          createMetricWithColumns("sales", "Sales", {
            p1: "sales_old",
            p2: "sales_new",
          }),
        ],
        comparisonMode: "endpoint",
      });
      const result = comparison.rows[0].metrics[0];

      assert.equal(result.valueA, 10);
      assert.equal(result.valueB, 14);
      assert.equal(result.delta, 4);
    });
  }

  function runTests() {
    const results = tests.map(function (item) {
      try {
        item.run();
        return {
          name: item.name,
          status: "pass",
          message: "OK",
        };
      } catch (error) {
        return {
          name: item.name,
          status: "fail",
          message: error && error.stack ? error.stack : String(error),
        };
      }
    });

    renderResults(results);
    return results;
  }

  function renderResults(results) {
    const list = document.getElementById("testList");
    const totalCount = document.getElementById("totalCount");
    const passedCount = document.getElementById("passedCount");
    const failedCount = document.getElementById("failedCount");
    const status = document.getElementById("testStatus");
    const passed = results.filter(function (result) {
      return result.status === "pass";
    }).length;
    const failed = results.length - passed;

    totalCount.textContent = String(results.length);
    passedCount.textContent = String(passed);
    failedCount.textContent = String(failed);
    status.textContent = failed ? "Есть ошибки: " + failed : "Все тесты прошли успешно.";

    list.innerHTML = "";
    results.forEach(function (result) {
      const item = document.createElement("article");
      const marker = document.createElement("span");
      const body = document.createElement("div");
      const name = document.createElement("div");
      const message = document.createElement("div");

      item.className = "test-item " + result.status;
      marker.className = "test-status";
      marker.textContent = result.status === "pass" ? "✓" : "!";
      name.className = "test-name";
      name.textContent = result.name;
      message.className = "test-message";
      message.textContent = result.message;

      body.append(name, message);
      item.append(marker, body);
      list.append(item);
    });
  }

  document.addEventListener("DOMContentLoaded", function () {
    registerTests();
    runTests();

    document.getElementById("runTestsButton").addEventListener("click", runTests);
  });
})(window);
