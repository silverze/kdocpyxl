/**
 * 这个 AirScript 脚本用于高效地读取或写入金山文档智能表格的指定单元格、一行或一列有效数据；
 * 以及获取指定 sheet 页面的有效行数和列数。
 */
 
// 从脚本参数中获取参数
const {
    sheet_name,
    row, column, value,
    is_column, values,
    get_rows_and_columns
  } = Context.argv;

  // 参数验证
  if (!sheet_name) {
    return { status: 'error', message: '缺少工作表名称参数' };
  }

  function validateNumericParams(...params) {
    return params.every(param => typeof param === 'number');
  }

  function operateSingleCell(sheet, row, column, value) {
    if (!validateNumericParams(row, column)) {
      return { status: 'error', message: 'row 和 column 必须是数字' };
    }

    const cell = sheet.Cells(row, column);
    if (value !== undefined) {
      cell.Value = value;
      return { status: 'success', message: '数据写入完成' };
    } else {
      return { status: 'success', message: cell.Text };
    }
  }

  function getLastNonEmptyCell(sheet, index, is_column) {
    let last = 1;
    let maxLast = 1;
    while (true) {
      const cell = is_column ? sheet.Cells(last, index) : sheet.Cells(index, last);
      if (cell.Value !== null && cell.Value !== undefined && cell.Value !== '') {
        maxLast = last;
      }
      last++;
      if (last > sheet.UsedRange.RowEnd && is_column) break;
      if (last > sheet.UsedRange.ColumnEnd && !is_column) break;
    }
    return maxLast;
  }

  function operateRowOrColumn(sheet, index, is_column, values) {
    console.log("Received params:", JSON.stringify({index, is_column, values}, null, 2));

    if (!validateNumericParams(index)) {
      return { status: 'error', message: 'index 必须是数字' };
    }

    if (values !== undefined) {
      if (Array.isArray(values)) {
        try {
          const range = is_column
            ? sheet.Range(sheet.Cells(1, index), sheet.Cells(values.length, index))
            : sheet.Range(sheet.Cells(index, 1), sheet.Cells(index, values.length));
          range.Value = is_column ? values.map(v => [v]) : [values];
          console.log("Data written successfully");
          return { status: 'success', message: '数据写入完成' };
        } catch (error) {
          console.error("Error writing data:", error);
          return { status: 'error', message: `写入数据时发生错误: ${error.message}` };
        }
      } else {
        return { status: 'error', message: 'values 必须是一个数组' };
      }
    } else {
      try {
        const lastNonEmpty = getLastNonEmptyCell(sheet, index, is_column);
        if (lastNonEmpty === 0) {
          return { status: 'success', message: [] };
        }
        const range = is_column
          ? sheet.Range(sheet.Cells(1, index), sheet.Cells(lastNonEmpty, index))
          : sheet.Range(sheet.Cells(index, 1), sheet.Cells(index, lastNonEmpty));
        const result = range.Value;
        return { status: 'success', message: JSON.stringify(result) };
      } catch (error) {
        console.error("Error reading data:", error);
        return { status: 'error', message: `读取数据时发生错误: ${error.message}` };
      }
    }
  }

  function getSheetRowsAndColumns(sheet) {
    try {
      const usedRange = sheet.UsedRange;
      const lastRow = usedRange.RowEnd;
      const lastColumn = usedRange.ColumnEnd;

      console.log(`原始行数和列数: 行数 = ${lastRow}, 列数 = ${lastColumn}`);

      // 可选：验证最后一行是否为空
      let actualLastRow = lastRow;
      while (actualLastRow > 0) {
        let rowIsEmpty = true;
        for (let col = 1; col <= lastColumn; col++) {
          if (sheet.Cells(actualLastRow, col).Value !== null && sheet.Cells(actualLastRow, col).Value !== '') {
            rowIsEmpty = false;
            break;
          }
        }
        if (!rowIsEmpty) break;
        actualLastRow--;
      }

      // 可选：验证最后一列是否为空
      let actualLastColumn = lastColumn;
      while (actualLastColumn > 0) {
        let columnIsEmpty = true;
        for (let row = 1; row <= actualLastRow; row++) {
          if (sheet.Cells(row, actualLastColumn).Value !== null && sheet.Cells(row, actualLastColumn).Value !== '') {
            columnIsEmpty = false;
            break;
          }
        }
        if (!columnIsEmpty) break;
        actualLastColumn--;
      }

      console.log(`调整后的行数和列数: 行数 = ${actualLastRow}, 列数 = ${actualLastColumn}`);
      return { status: 'success', message: {"Rows": actualLastRow, "Columns": actualLastColumn}};
    } catch (error) {
      console.error("获取表格行数和列数时发生错误:", error);
      return { status: 'error', message: `获取表格行数和列数时发生错误: ${error.message}` };
    }
  }

  try {
    const sheet = Application.Sheets.Item(sheet_name);
    sheet.Activate();

    if (get_rows_and_columns) {
      return getSheetRowsAndColumns(sheet);
    } else if (row !== undefined && column !== undefined) {
      return operateSingleCell(sheet, row, column, value);
    } else if (row !== undefined || column !== undefined) {
      const index = row !== undefined ? row : column;
      return operateRowOrColumn(sheet, index, column !== undefined, values);
    } else {
      return { status: 'error', message: '参数不足，无法执行操作' };
    }

  } catch (error) {
    console.error('操作失败:', error);
    return { status: 'error', message: `操作失败: ${error.message}` };
  }