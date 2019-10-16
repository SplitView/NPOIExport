using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;

namespace NPOI.Extensions.Export
{
    public class ExcelTable<T>
    {
        public List<T> Model { get; set; }

        private Dictionary<int, Expression<Func<T, object>>> _tableExpressions = new Dictionary<int, Expression<Func<T, object>>>();
        public int ActiveRow { get; set; }
        private readonly IWorkbook _workbook;
        private readonly ISheet _sheet;
        public ExcelTable()
        {
            _workbook = new XSSFWorkbook();
            _sheet = _workbook.CreateSheet("Sheet 1");
            ActiveRow = 0;

        }

        public void SetColumn(Expression<Func<T, object>> expression, int position)
        {
            _tableExpressions.Add(position, expression);
        }

        private IRow GetRow(T model)
        {
            var row = _sheet.GetRow(ActiveRow);

            foreach (var item in _tableExpressions.OrderBy(x => x.Key))
            {
                var cell = row.CreateCell(item.Key);
                var cellvalue = item.Value.Compile().Invoke(model);
                cell.SetCellValue(cellvalue);
            }
        }

        public IWorkbook GenerateWorkBook()
        {
            foreach (var item in Model)
            {
                var row = _sheet.CreateRow(ActiveRow);



                ActiveRow++;
            }

            return _workbook;
        }
    }
    public class TestClass
    {
        public void TestMethod()
        {
            var excelTable = new ExcelTable<Student>();

            excelTable.SetColumn(x => x.Name, 0);
            excelTable.SetColumn(x => x.Roll, 1);
        }
    }

    public class Student
    {
        public string Name { get; set; }
        public int Roll { get; set; }
    }
}
