﻿using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Application = Microsoft.Office.Interop.Excel.Application;
using Path = System.IO.Path;
using Shape = Microsoft.Office.Interop.Excel.Shape;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;
using XlHAlign = Microsoft.Office.Interop.Excel.XlHAlign;

namespace EstimatesAssembly {
    class BookEstimates {
        private const string pageContent = @"\contentpage.xlsx";
        private const string pageTitle = @"\titlepage.xlsx";
        private const string pageResolution = @"\resolutionpage.xlsx";
        private const int PixelW = 75;
        private const int PixelH = 38;
        public ProgressBar _pgBar;

        struct Ogl {
            public string col1;
            public string col2;
        }

        public ProgressBar PgBar {
            get { return _pgBar; }
            set { _pgBar = value; }
        }

        private string _nameBook;
        private string _pathBook;
        public Application Ex;
        public Workbook Wb;
        public Workbook TmpWb;
        // Для перенумерации листов книги
        private const int stopPos = 49;
        private const int endPos = 59;
        private static int delta = 0;


        public string NameBook {
            get { return _nameBook; }
            set { _nameBook = value; }
        }

        public string PathBook {
            get { return _pathBook; }
            set { _pathBook = value; }
        }

        public BookEstimates() {
            Ex = new Application { Visible = false, DisplayAlerts = false };
        }

        public void ShowExcel(Boolean show) {
            Ex.Visible = show;
        }
        // Тип сметы
        public static int FindTypeSheet(Worksheet sheet) {
            if (sheet.Range["A1", "Q15"].Find( @"ЛОКАЛЬНЫЙ СМЕТНЫЙ РАСЧЕТ №" ) != null) {
                return 1;
            } else if (sheet.Range["A1", "Q15"].Find( @"ОБЪЕКТНЫЙ СМЕТНЫЙ РАСЧЕТ №" ) != null) {
                return 2;
            } else if (sheet.Range["A1", "Q15"].Find( @"ВЕДОМОСТЬ РЕСУРСОВ" ) != null) {
                return 3;
            } else if (sheet.Range["A1", "Q15"].Find( @"ВЕДОМОСТЬ ОБЪЕМОВ РАБОТ №" ) != null) {
                return 4;
            } else if (sheet.Range["A1", "Q15"].Find( @"СВОДНЫЙ СМЕТНЫЙ РАСЧЕТ СТОИМОСТИ СТРОИТЕЛЬСТВА" ) != null) {
                return 5;
            } else if (sheet.Range["A1", "Q15"].Find( @"Локальный ресурсный сметный расчет" ) != null) {
                return 6;
            } else if (sheet.Range["A1", "Q15"].Find( @"СВОДНАЯ ВЕДОМОСТЬ РЕСУРСОВ" ) != null) {
                return 7;
            }
            return 0;
        }

        // Добавить элемент(ы) в книгу
        public void AddSheetNew(string[] selectedItems) {
            if (selectedItems.Length == 0) {
                MessageBox.Show( @"Не выбрано ни одной сметы!", @"Внимание!" );
                return;
            }
            Wb = Ex.Workbooks.Count == 0 ? Ex.Workbooks.Add() : Ex.ActiveWorkbook;
            _pgBar.Maximum = selectedItems.Length + 2;
            _pgBar.Minimum = 0;
            _pgBar.Value = 0;
            Boolean isBookEstimate = false;
            string tmpfile = null;
            foreach (string selectedItem in selectedItems) {
                TmpWb = Ex.Workbooks.Open( selectedItem );
                Worksheet title = TmpWb.Sheets[1];
                if (title.Name.Equals( @"Титул" )) {
                    isBookEstimate = true;
                    tmpfile = selectedItem;
                    break;
                }
            }
            if (isBookEstimate) {
                TmpWb = Ex.Workbooks.Open( tmpfile );
                foreach (Worksheet sheet in TmpWb.Sheets) {
                    sheet.Copy( Type.Missing, Wb.ActiveSheet );
                    _pgBar.Value++;
                }
                TmpWb.Close();
            } else {
                foreach (string selectedItem in selectedItems) {
                    TmpWb = Ex.Workbooks.Open( selectedItem );
                    foreach (Worksheet sheet in TmpWb.Sheets) {
                        switch (FindTypeSheet( sheet )) {
                            case 1:
                                WorkWithExcelLs( sheet, selectedItem );
                                break;
                            case 2:
                                WorkWithExcelOs( sheet, selectedItem );
                                break;
                            case 3:
                                WorkWithExcelR( sheet );
                                break;
                            case 4:
                                WorkWithExcelVR( sheet );
                                break;
                            case 5:
                                WorkWithExcelSSR( sheet );
                                break;
                            case 6:
                                WorkWithExcelLRS( sheet );
                                break;
                        }
                        sheet.Copy( Type.Missing, Wb.ActiveSheet );
                    }
                    TmpWb.Close();
                    _pgBar.Value++;
                }
            }
            _pgBar.Value = 0;
            foreach (string myvar in GetListSheet()) {
                if (myvar.Contains( "Лист" )) {
                    Wb.Sheets[myvar].Delete();
                }
            }
        }

        // Обработка сводного сметного расчета
        private void WorkWithExcelSSR(Worksheet sheet) {
            Range find = sheet.Cells.Find( "Итого \"Налоги и обязательные платежи\"" );
            if (find != null) {
                sheet.Name = @"СС 01";
            } else {
                sheet.Name = @"СС 02";
            }
        }

        // Обработка локального ресурсного сметного расчета
        private void WorkWithExcelLRS(Worksheet sheet) {
            Range find = sheet.Cells.Find( "к Локальной смете №" );
            string number = find.Value2;
            number = number.Substring( number.IndexOf( "№" ) + 2 );
            sheet.Name = "ЛР" + number;
        }

        // Удалить элемент(ы) из книги
        public void DeleteSheet(ListView.SelectedListViewItemCollection selectedItems, ref ProgressBar pgBar) {
            if (selectedItems.Count == 0) {
                MessageBox.Show( @"Не выбрано ни одно сметы!", @"Внимание!" );
                return;
            }
            if (Ex.Workbooks.Count == 0) {
                return;
            }
            Ex.DisplayAlerts = false;
            pgBar.Maximum = Wb.Sheets.Count;
            pgBar.Minimum = 1;
            pgBar.Value = 1;
            foreach (ListViewItem selectedItem in selectedItems) {
                Worksheet worksheet = Wb.Sheets[selectedItem.Text];
                if (worksheet.Visible == XlSheetVisibility.xlSheetHidden) {
                    worksheet.Visible = XlSheetVisibility.xlSheetVisible;
                }
                if (Wb.Sheets.Count == 1) {
                    Wb.Sheets.Add();
                }
                pgBar.PerformStep();
                Wb.Sheets[selectedItem.Text].Delete();
            }
            pgBar.Value = 1;
        }

        // Возвращает список листов в книге
        public IEnumerable<string> GetListSheet() {
            var list = new List<string>();
            if (Ex.Workbooks.Count == 0) {
                return null;
            }
            Workbook workbook = Ex.ActiveWorkbook;
            if (workbook.Sheets.Count == 0) {
                return null;
            }
            for (int i = 1; i < workbook.Sheets.Count + 1; i++) {
                list.Add( workbook.Sheets[i].Name );
            }
            return list;
        }

        // Сохранение тома
        public void SaveWorkbook() {
            string fullname = Path.Combine( _pathBook, _nameBook + @".xls" );
            if (File.Exists( fullname )) {
                DialogResult dlgres = MessageBox.Show( @"Книга уже существует. Переписать?", @"Внимание!",
                    MessageBoxButtons.OKCancel );
                if (dlgres == DialogResult.Cancel) {
                    return;
                }
            }
            Ex.DisplayAlerts = false;
            Ex.UserControl = true;
            try {
                Ex.ActiveWorkbook.SaveAs( fullname,
                    XlFileFormat.xlWorkbookNormal, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing );
                MessageBox.Show( @"Книга успешно сохранена!" );
            } catch (Exception e) {
                MessageBox.Show( e.Message );
            }
        }

        // Закрытие рабочего Excel-приложения
        public void CloseBook() {
            Ex.DisplayAlerts = false;
            Ex.UserControl = true;
            Ex.Quit();
        }

        // Инициализация книги
        public void initBook(string bookfile) {
            if (File.Exists( bookfile )) {
                Wb = Ex.Workbooks.Open( bookfile );
                Ex.DisplayAlerts = false;
                Ex.UserControl = true;
            }
        }

        // Перемещение элемента вверх по списку
        public void MoveUpsheet(ListView.SelectedListViewItemCollection selectedItems) {
            if (selectedItems.Count == 0) {
                MessageBox.Show( @"Не выбрано ни одно сметы!", @"Внимание!" );
                return;
            }
            if (Ex.Workbooks.Count == 0) {
                return;
            }
            foreach (ListViewItem selectedItem in selectedItems) {
                Worksheet worksheet = Wb.Sheets[selectedItem.Text];
                int i = worksheet.Index;
                if (i > 1) {
                    worksheet.Move( Wb.Sheets[i - 1], Type.Missing );
                }
            }
        }

        // Перемещение элемента вниз по списку
        public void MoveDownsheet(ListView.SelectedListViewItemCollection selectedItems) {
            if (selectedItems.Count == 0) {
                MessageBox.Show( @"Не выбрано ни одно сметы!", @"Внимание!" );
                return;
            }
            if (Ex.Workbooks.Count == 0) {
                return;
            }
            foreach (ListViewItem selectedItem in selectedItems) {
                Worksheet worksheet = Wb.Sheets[selectedItem.Text];
                int i = worksheet.Index;
                if (i < Wb.Sheets.Count) {
                    worksheet.Move( Type.Missing, Wb.Sheets[i + 1] );
                }
            }
        }

        // Сортировка списка согласно правилам сметчиков
        public void SortWorksheets(ref ProgressBar pgBar) {
            List<string> list = new List<string>();
            if (Ex.ActiveWorkbook == null) {
                return;
            }
            foreach (Worksheet ws in Ex.ActiveWorkbook.Sheets) {
                list.Add( ws.Name );
            }
            list.Sort( Compare );
            Workbook wb = Ex.ActiveWorkbook;
            pgBar.Maximum = list.Count;
            pgBar.Minimum = 1;
            foreach (string str in list) {
                Worksheet ws = wb.Sheets[str];
                ws.Move( Wb.Sheets[list.IndexOf( str ) + 1], Type.Missing );
                pgBar.PerformStep();
            }
            pgBar.Value = 1;
        }

        // Компаратор для сортировщика
        public int Compare(String x, String y) {
            Regex pattern = new Regex( @"(?:[ЛОРС]+?[СР]?)*(?<ss>(\(?(\d{2,4})[-|\.])*(\d{1,4})(\)?))" );
            MatchCollection mc1 = pattern.Matches( x );
            MatchCollection mc2 = pattern.Matches( y );

            if (mc1.Count == 0) {
                x = @"О0";
            }
            if (mc2.Count == 0) {
                y = @"О0";
            }
            int compareResult = 0;
            Int64 xx1;
            Int64 yy1;
            int xi = x.IndexOf( "." );
            int yi = y.IndexOf( "." );
            if (xi > 0) {
                x = x.Remove( xi );
            }
            if (yi > 0) {
                y = y.Remove( yi );
            }
            if (TwoChar( x )) {
                xx1 = Int64.Parse( x.Substring( 2 ).Replace( "-", "" ).Replace( ".", "" ).Replace( "(", "" ).Replace( " ", "" ).Replace( ")", "" ).PadRight( 14, '0' ) );
            } else {
                xx1 = Int64.Parse( x.Substring( 1 ).Replace( "-", "" ).Replace( ".", "" ).Replace( "(", "" ).Replace( " ", "" ).Replace( ")", "" ).PadRight( 14, '0' ) );
            }

            if (TwoChar( y )) {
                yy1 = Int64.Parse( y.Substring( 2 ).Replace( "-", "" ).Replace( ".", "" ).Replace( "(", "" ).Replace( " ", "" ).Replace( ")", "" ).PadRight( 14, '0' ) );
            } else {
                yy1 = Int64.Parse( y.Substring( 1 ).Replace( "-", "" ).Replace( ".", "" ).Replace( "(", "" ).Replace( " ", "" ).Replace( ")", "" ).PadRight( 14, '0' ) );
            }

            Int64 xx2 = ConvertChar( x );
            Int64 yy2 = ConvertChar( y );

            if (xx1 > yy1) {
                compareResult = 1;
            } else if (xx1 < yy1) {
                compareResult = -1;
            } else if (xx1 == yy1) {
                if (xx2 > yy2) {
                    compareResult = 1;
                } else if (xx2 < yy2) {
                    compareResult = -1;
                } else {
                    compareResult = 0;
                }
            }
            return compareResult;
        }

        // В названии первых символов - два?
        private bool TwoChar(string s) {
            if (s.Contains( "С" )) {
                return true;
            }
            return false;
        }

        // Выдает число в зависимости от первых двух символов наименования
        private Int64 ConvertChar(string a) {
            switch (a.Substring( 0, 1 )) {
                case "СС":
                case "ОС":
                case "О":
                    return 1;
                case "ЛC":
                case "Л":
                    return 2;
                case "Р":
                    return 3;
                default:
                    return 0;
            }
        }

        // Дополнительная обработка таблиц
        public void AdaptionSheets(ref ProgressBar pgBar) {
            Workbook mainBook = Ex.ActiveWorkbook;
            Range r;
            if (mainBook == null) {
                return;
            }
            pgBar.Maximum = mainBook.Sheets.Count;
            pgBar.Minimum = 1;
            foreach (Worksheet worksheet in mainBook.Sheets) {
                if (!worksheet.Name.Equals( @"Титул" )
                    && !worksheet.Name.Equals( "@Огл" )
                    && !worksheet.Name.Equals( @"Разрешение" )) {
                    worksheet.Activate();
                    Range rr = worksheet.Range["A1", "A1"];
                    rr.Activate();
                    HPageBreaks hbreak = worksheet.HPageBreaks;
                    int pageCount = worksheet.HPageBreaks.Count;
                    if (pageCount != 0) {
                        r = hbreak.Item[pageCount].Location;
                        int t = FindLastRow( worksheet );
                        if (( t - r.Row ) < 12) {
                            var tmpr = worksheet.Range["A" + Convert.ToString( r.Row - 12 )];
                            hbreak.Item[pageCount].Location = tmpr;
                        }
                    }
                }
                pgBar.PerformStep();
            }
            pgBar.Value = 1;
        }

        private int FindLastRow(Worksheet worksheet) {
            Range r = worksheet.Cells[worksheet.Rows.Count, 1];
            return r.get_End( XlDirection.xlUp ).Row;
        }

        public void NumberingPage(ref ProgressBar pgBar) {
            //Ex.ScreenUpdating = false;
            Workbook mainBook = Ex.ActiveWorkbook;
            if (mainBook == null) {
                return;
            }
            pgBar.Maximum = mainBook.Sheets.Count;
            pgBar.Minimum = 1;
            foreach (Worksheet worksheet in mainBook.Sheets) {
                if (worksheet.Name.Contains( @"Титул" ))
                    worksheet.Delete();
                else
                    if (worksheet.Name.Contains( @"Разрешение" ))
                    worksheet.Delete();
                else
                        if (worksheet.Name.Contains( @"Огл" ))
                    worksheet.Delete();
                else
                            if (worksheet.Name.Contains( @"Лист" ))
                    worksheet.Delete();
                pgBar.PerformStep();
            }
            // Вставим оглавление
            Workbook tmpContent = Ex.Workbooks.Open( MainFormAsm.iniSet.TxtToolsFilesPath + pageContent );
            tmpContent.Worksheets[1].Copy( mainBook.Sheets[1], Type.Missing );
            tmpContent.Close();
            Worksheet ogl = mainBook.Sheets[1];
            //            Worksheet title = mainBook.Sheets[1];
            ogl.Name = @"Огл";
            ogl.Cells[2, 5] = _nameBook;
            // временно
            Worksheet title = mainBook.Sheets.Add( mainBook.Sheets.Item[1], Type.Missing, 1, XlSheetType.xlWorksheet );
            title.Name = @"Титул";
            //title.Cells[1, 1] = @"1";
            pgBar.Value = 1;
            foreach (Worksheet worksheet in mainBook.Sheets) {
                worksheet.Select();
                Ex.ActiveWindow.View = XlWindowView.xlPageBreakPreview;
                pgBar.PerformStep();
            }
            pgBar.Value = 1;
            foreach (Worksheet worksheet in mainBook.Sheets) {
                worksheet.Select();
                worksheet.PageSetup.FitToPagesWide = 1;
                worksheet.PageSetup.FitToPagesTall = 999;
                worksheet.PageSetup.Zoom = false;
                if (worksheet.VPageBreaks.Count > 0) {
                    worksheet.VPageBreaks.get_Item( 1 ).DragOff( XlDirection.xlToRight, 1 );
                }
                pgBar.PerformStep();
            }
            pgBar.Value = 1;
            int ns = 3;
            int x = 1;
            Ogl a = new Ogl();
            if (mainBook.Sheets.Count < stopPos - 1) {
                Range clr = ogl.Range["A60", "L125"];
                clr.Clear();
            }
            foreach (Worksheet worksheet in mainBook.Sheets) {
                worksheet.Select();
                worksheet.PageSetup.FirstPageNumber = x;
                worksheet.PageSetup.RightFooter = "&P";
                worksheet.PageSetup.LeftHeader = " ";
                worksheet.PageSetup.CenterHeader = " ";
                worksheet.PageSetup.RightHeader = " ";
                a = GetColumnsSheet( worksheet );
                if (!worksheet.Name.Equals( "Титул" ) && !worksheet.Name.Equals( "Огл" )) {
                    ogl.Cells[ns, 4] = ns - delta - 2;
                    ogl.Cells[ns, 5] = a.col1;
                    ogl.Cells[ns, 8] = a.col2;
                    Range range_1 = ogl.Cells[ns, 12];
                    range_1.Value2 = String.Format( "{0}", worksheet.PageSetup.FirstPageNumber );
                    ogl.Hyperlinks.Add( range_1, "", "'" + worksheet.Name + "'!A1", Type.Missing, "Hyperlink Test" );
                    Range ssss = ogl.Rows[ns];
                    ssss.RowHeight = 12.75;
                    ns++;
                }
                if (ns == stopPos + 1) {
                    ns = endPos + 1;
                    delta = 10;
                }
                x = worksheet.PageSetup.FirstPageNumber + worksheet.PageSetup.Pages.Count;
                pgBar.PerformStep();
            }
            delta = 0;
            //ogl.Range["D3", "L4"].Clear();
            title.Delete();
            // Вставим титульные листы
            Workbook tmpTitle = Ex.Workbooks.Open( MainFormAsm.iniSet.TxtToolsFilesPath + pageTitle );
            tmpTitle.Worksheets[1].Copy( mainBook.Sheets[1], Type.Missing );
            tmpTitle.Close();
            Worksheet titles = mainBook.Sheets[1];
            titles.Name = @"Титул";
            TitleFill( ref titles );
            //AdaptionSheets(ref pgBar);
            if (int.Parse( MainFormAsm.iniSet.NumModification ) != 0) {
                Workbook tmpResolution = Ex.Workbooks.Open( MainFormAsm.iniSet.TxtToolsFilesPath + pageResolution );
                tmpResolution.Worksheets[1].Copy( mainBook.Sheets[2], Type.Missing );
                tmpResolution.Close();
                Worksheet resolution = mainBook.Sheets[2];
                resolution.Name = @"Разрешение";
                ResolutionFill( ref resolution );
            }
            if (ns < endPos + 1) {
                StampFill( false, ref ogl, x - 2 );
            } else {
                StampFill( true, ref ogl, x - 2 );
            }
            pgBar.Value = 1;
        }

        // Заполним Разрешение
        private void ResolutionFill(ref Worksheet resolution) {
            resolution.Cells[9, 5] = MainFormAsm.iniSet.NumModification;
            resolution.Cells[9, 7] = MainFormAsm.iniSet.TbPageNumber;
            resolution.Cells[46, 8] = MainFormAsm.iniSet.TbChiefEngineer;
            resolution.Cells[47, 8] = MainFormAsm.iniSet.CbGipText;
            resolution.Cells[48, 8] = MainFormAsm.iniSet.CbMadeInText;
            resolution.Cells[49, 8] = MainFormAsm.iniSet.CbMadeInText;
            resolution.Cells[46, 13] = MainFormAsm.iniSet.DateAjustment.ToString( "MM.yy", CultureInfo.CreateSpecificCulture( "ru-RU" ) );
            resolution.Cells[47, 13] = MainFormAsm.iniSet.DateAjustment.ToString( "MM.yy", CultureInfo.CreateSpecificCulture( "ru-RU" ) );
            resolution.Cells[48, 13] = MainFormAsm.iniSet.DateAjustment.ToString( "MM.yy", CultureInfo.CreateSpecificCulture( "ru-RU" ) );
            resolution.Cells[49, 13] = MainFormAsm.iniSet.DateAjustment.ToString( "MM.yy", CultureInfo.CreateSpecificCulture( "ru-RU" ) );
            InsertImage( ref resolution, 47, 11, MainFormAsm.iniSet.TbChiefEngineer );
            InsertImage( ref resolution, 48, 11, MainFormAsm.iniSet.CbGipText );
            InsertImage( ref resolution, 49, 11, MainFormAsm.iniSet.CbMadeInText );
            InsertImage( ref resolution, 50, 11, MainFormAsm.iniSet.CbMadeInText );
            resolution.Cells[46, 15] = "ООО \"Технологии проектирования\"";
            resolution.Cells[48, 23] = "1";
            resolution.Cells[3, 5] = MainFormAsm.iniSet.TbDocumentNumber;
            // 
            String loverStr = MainFormAsm.iniSet.ListTypeDocument.ToLower();
            String volNum = MainFormAsm.iniSet.NumVolumeNumber;
            String bookNum = MainFormAsm.iniSet.NumBookNumber;
            String partNum = MainFormAsm.iniSet.NumPartNumber;
            String lStr = loverStr.Substring( 0, 1 ).ToUpper() + loverStr.Substring( 1, loverStr.Length - 1 );
            String str = @"Инв.№" + MainFormAsm.iniSet.TbInventoryNumber + "\n" +
                MainFormAsm.iniSet.TbCodeObject + "\n" +
                @"Том " + volNum + "." + bookNum + "." + partNum + " \"" + lStr + "\"";
            resolution.Cells[1, 9] = str;
            //
            resolution.Cells[1, 19] = MainFormAsm.iniSet.TbNameObject;
        }

        // Заполним титул
        private void TitleFill(ref Worksheet title) {

            title.Cells[8, 3] = MainFormAsm.iniSet.TbCertificate; // Свидетельство
            title.Cells[57, 3] = MainFormAsm.iniSet.TbCertificate;
            title.Cells[10, 3] = MainFormAsm.iniSet.TbCustomer; // Заказчик
            title.Cells[59, 3] = MainFormAsm.iniSet.TbCustomer;
            title.Cells[13, 3] = MainFormAsm.iniSet.TbNameObject;
            title.Cells[61, 3] = MainFormAsm.iniSet.TbNameObject;
            if (!MainFormAsm.iniSet.CbRebuild) {
                title.Cells[22, 3] = "РАЗДЕЛ " + int.Parse( MainFormAsm.iniSet.NumVolumeNumber ) + " \"СМЕТА НА СТРОИТЕЛЬСТВО\"";
                title.Cells[70, 3] = "РАЗДЕЛ " + int.Parse( MainFormAsm.iniSet.NumVolumeNumber ) + " \"СМЕТА НА СТРОИТЕЛЬСТВО\"";
            } else {
                title.Cells[22, 3] = "РАЗДЕЛ " + int.Parse( MainFormAsm.iniSet.NumVolumeNumber ) + " \"СМЕТА НА КАПИТАЛЬНЫЙ РЕМОНТ\"";
                title.Cells[70, 3] = "РАЗДЕЛ " + int.Parse( MainFormAsm.iniSet.NumVolumeNumber ) + " \"СМЕТА НА КАПИТАЛЬНЫЙ РЕМОНТ\"";
            }

            title.Cells[24, 3] = @"ЧАСТЬ 2 " + MainFormAsm.iniSet.ListTypeDocument.ToUpper();
            title.Cells[72, 3] = @"ЧАСТЬ 2 " + MainFormAsm.iniSet.ListTypeDocument.ToUpper();

            title.Cells[25, 3] = "КНИГА " + MainFormAsm.iniSet.NumBookNumber;
            title.Cells[73, 3] = "КНИГА " + MainFormAsm.iniSet.NumBookNumber;

            title.Cells[27, 3] = MainFormAsm.iniSet.TbCodeObject;
            title.Cells[75, 3] = MainFormAsm.iniSet.TbCodeObject;
            string sss = "ТОМ " + MainFormAsm.iniSet.NumVolumeNumber + "." +
                         MainFormAsm.iniSet.NumBookNumber + "." +
                         MainFormAsm.iniSet.NumPartNumber;
            title.Cells[29, 3] = sss;
            title.Cells[77, 3] = sss;

            title.Cells[79, 10] = MainFormAsm.iniSet.TbChiefEngineer.ToUpper();
            InsertImage( ref title, 79, 8, MainFormAsm.iniSet.TbChiefEngineer.ToUpper() );
            title.Cells[81, 10] = MainFormAsm.iniSet.CbGipText.ToUpper();
            InsertImage( ref title, 81, 8, MainFormAsm.iniSet.CbGipText.ToUpper() );

            title.Cells[49, 3] = MainFormAsm.iniSet.TbYearTitle; // Год 
            title.Cells[96, 3] = MainFormAsm.iniSet.TbYearTitle;

            title.Cells[92, 2] = MainFormAsm.iniSet.TbInventoryNumber;

            switch (int.Parse( MainFormAsm.iniSet.NumModification )) {
                case 0:
                    title.Range["D38", "H39"].UnMerge();
                    title.Range["D38", "H39"].Clear();
                    title.Range["D87", "H88"].UnMerge();
                    title.Range["D87", "H88"].Clear();
                    break;
                default:
                    title.Cells[39, 4] = int.Parse( MainFormAsm.iniSet.NumModification );
                    title.Cells[39, 5] = MainFormAsm.iniSet.TbDocumentNumber;
                    InsertImage( ref title, 39, 6, MainFormAsm.iniSet.CbMadeInText );
                    title.Cells[39, 7] = MainFormAsm.iniSet.DateAjustment.ToString( "\tMM/yyyy" );
                    title.Cells[88, 4] = int.Parse( MainFormAsm.iniSet.NumModification );
                    title.Cells[88, 5] = MainFormAsm.iniSet.TbDocumentNumber;
                    InsertImage( ref title, 88, 6, MainFormAsm.iniSet.CbMadeInText );
                    title.Cells[88, 7] = MainFormAsm.iniSet.DateAjustment.ToString( "\tMM/yyyy" );
                    break;
            }
        }

        // Заполним штамп оглавления
        private void StampFill(Boolean twoPage, ref Worksheet stamp, int x) {
            stamp.Cells[endPos - 4, 5] = MainFormAsm.iniSet.CbMadeInText;
            InsertImage( ref stamp, endPos - 5, 7, MainFormAsm.iniSet.CbMadeInText );
            stamp.Cells[endPos - 2, 5] = MainFormAsm.iniSet.TbHeadDepartment;
            InsertImage( ref stamp, endPos - 3, 7, MainFormAsm.iniSet.TbHeadDepartment );
            stamp.Cells[endPos - 1, 5] = MainFormAsm.iniSet.CbGipText;
            InsertImage( ref stamp, endPos - 2, 7, MainFormAsm.iniSet.CbGipText );
            stamp.Cells[endPos - 4, 8] = MainFormAsm.iniSet.DateToStamp.ToString( "MM.yy" );
            stamp.Cells[endPos - 2, 8] = MainFormAsm.iniSet.DateToStamp.ToString( "MM.yy" );
            stamp.Cells[endPos - 1, 8] = MainFormAsm.iniSet.DateToStamp.ToString( "MM.yy" );
            stamp.Cells[endPos - 7, 9] = MainFormAsm.iniSet.TbCodeObject;
            stamp.Cells[endPos - 5, 9] = MainFormAsm.iniSet.TbNameObject + "\nОбъектные и локальные сметы";
            stamp.Cells[endPos - 2, 10] = "ООО \"Тезнологии проектирования\"";
            stamp.Cells[endPos - 4, 10] = MainFormAsm.iniSet.CbStageDevelope.Substring( 0, 1 );
            stamp.Cells[endPos - 4, 11] = "1";
            stamp.Cells[endPos - 4, 12] = ( x + 1 ).ToString( CultureInfo.InvariantCulture );
            stamp.Cells[endPos - 4, 2] = MainFormAsm.iniSet.TbInventoryNumber;
            if (int.Parse( MainFormAsm.iniSet.NumModification ) != 0) {
                stamp.Cells[endPos - 6, 3] = MainFormAsm.iniSet.NumModification;
                stamp.Cells[endPos - 6, 4] = "-";
                stamp.Cells[endPos - 6, 5] = MainFormAsm.iniSet.TbPageNumber;
                stamp.Cells[endPos - 6, 6] = MainFormAsm.iniSet.TbDocumentNumber;
                stamp.Cells[endPos - 6, 8] = MainFormAsm.iniSet.DateAjustment.ToString( "MM.yy" );
                InsertImage( ref stamp, endPos - 7, 7, MainFormAsm.iniSet.CbMadeInText );
            }

            if (twoPage) {
                stamp.Cells[endPos + 55, 2] = MainFormAsm.iniSet.TbInventoryNumber;
                stamp.Cells[endPos + 57, 9] = MainFormAsm.iniSet.TbCodeObject;
            }
        }

        // Вставить картинку
        private void InsertImage(ref Worksheet sheet, int y, int x, string fio) {
            char[] charsToTrim = { '\n', '\r', ' ' };
            string imgFile;
            Shape shape = null;
            Range range = sheet.Cells[y, x];
            fio = fio.TrimEnd( charsToTrim );
            float xx = (float) ( (double) range.Left - 10 );
            float yy = (float) ( (double) range.Top - 20 );
            try {
                var fName1 = MainFormAsm.iniSet.TxtImagePath + @"\" + ConvertName( fio ) + ".jpg";
                var fName2 = MainFormAsm.iniSet.TxtImagePath + @"\" + ConvertName( fio ) + ".tif";
                var fName3 = MainFormAsm.iniSet.TxtImagePath + @"\" + ConvertName( fio ) + ".tiff";
                var fName4 = MainFormAsm.iniSet.TxtImagePath + @"\" + ConvertName( fio ) + ".png";
                if (File.Exists( fName1 )) {
                    imgFile = fName1;
                } else if (File.Exists( fName2 )) {
                    imgFile = fName2;
                } else if (File.Exists( fName3 )) {
                    imgFile = fName3;
                } else if (File.Exists( fName4 )) {
                    imgFile = fName4;
                } else {
                    return;
                }
                shape = sheet.Shapes.AddPicture( imgFile, MsoTriState.msoTrue, MsoTriState.msoTrue,
                    xx, yy, PixelW, PixelH );
                if (shape != null) {
                    shape.PictureFormat.TransparentBackground = MsoTriState.msoTrue;
                    shape.PictureFormat.TransparencyColor = ColorTranslator.ToOle( Color.White );
                    shape.Fill.Visible = MsoTriState.msoFalse;
                }
            } catch (Exception e) {
                MessageBox.Show( e.Message, @"Ошибка при работе с изображением!" );
            }
        }

        //// Вставить картинку
        //private void InsertImage(ref Worksheet sheet, int y, int x, string fio) {
        //    Range range = sheet.Cells[y, x];
        //    float xx = FloatLeftPixelsCalculation( range );
        //    float yy = FloatTopPixelsCalculation( range );
        //    try {
        //        var fName1 = MainFormAsm.iniSet.TxtImagePath + @"\" + ConvertName( fio ) + ".jpg";
        //        var fName2 = MainFormAsm.iniSet.TxtImagePath + @"\" + ConvertName( fio ) + ".tif";
        //        var fName3 = MainFormAsm.iniSet.TxtImagePath + @"\" + ConvertName( fio ) + ".png";
        //        if (File.Exists( fName1 )) {
        //            Shape shape = sheet.Shapes.AddPicture( fName1, MsoTriState.msoTrue, MsoTriState.msoTrue, xx, yy, PixelW, PixelH );
        //        } else if (File.Exists( fName2 )) {
        //            Shape shape = sheet.Shapes.AddPicture( fName2, MsoTriState.msoTrue, MsoTriState.msoTrue, xx, yy, PixelW, PixelH );
        //        } else if (File.Exists( fName3 )) {
        //            Shape shape = sheet.Shapes.AddPicture( fName3, MsoTriState.msoTrue, MsoTriState.msoTrue, xx, yy, PixelW, PixelH );
        //        }

        //    } catch (Exception e) {
        //        MessageBox.Show( e.Message, @"Ошибка при работе с изображением!" );
        //    }
        //}

        private static string ConvertName(string name) {
            string n = name.Replace( ".", "_" ).Replace( " ", "_" );
            n = n.Substring( 0, n.Length - 1 );
            return n;
        }

        // Вытаскиваем из таблицы номер и наименование сметы или объекта
        private Ogl GetColumnsSheet(_Worksheet worksheet) {
            Ogl o = new Ogl();
            Range range = worksheet.Cells.Find( @"ЛОКАЛЬНЫЙ СМЕТНЫЙ РАСЧЕТ" );
            if (range != null) {
                o.col1 = @"Лок.см. " + worksheet.Name.Substring( 2 );
                o.col2 = worksheet.Range["D12"].Value2;
                return o;
            }
            range = worksheet.Cells.Find( @"ОБЪЕКТНЫЙ СМЕТНЫЙ РАСЧЕТ" );
            if (range != null) {
                o.col1 = @"Об.см. " + worksheet.Name.Substring( 2 );
                o.col2 = worksheet.Range["D8"].Value2;
                return o;
            }
            range = worksheet.Cells.Find( @"ВЕДОМОСТЬ РЕСУРСОВ" );
            if (range != null) {
                o.col1 = @"Рес.вед. " + worksheet.Name.Substring( 1 );
                o.col2 = worksheet.Range["C8"].Value2;
                return o;
            }
            return o;
        }

        public void RebuildWorksheets(ref ProgressBar pgBar) {
            Workbook mainBook = Ex.ActiveWorkbook;
            pgBar.Maximum = mainBook.Sheets.Count;
            pgBar.Minimum = 1;
            foreach (Worksheet worksheet in mainBook.Sheets) {
                string sss = WorkWithExcelLs( worksheet, worksheet.Name );
                pgBar.PerformStep();
            }
        }

        private string QuarterFromDate(DateTime value) {
            int a = DateAndTime.DatePart( DateInterval.Quarter, value );
            int b = DateAndTime.DatePart( DateInterval.Year, value );
            if (MainFormAsm.iniSet.CbQuarter) {
                return String.Format( "{0}-й квартал {1} года.", a, b );
            } else {
                return value.ToString( "dd MMMM yyyy", CultureInfo.CreateSpecificCulture( "ru-RU" ) );
            }
        }

        public static string RenameName(string name) {
            Regex pattern = new Regex( @"(?:\D*)(?<ss>((\d{2,4})[-|\.])*(\d{1,4}))" );
            MatchCollection mc = pattern.Matches( name );
            if (mc.Count > 0) {
                GroupCollection groups = mc[0].Groups;
                return groups["ss"].Value;
            } else {
                return name;
            }
        }
        public static string RemoveBeginPos(string name) {
            Regex pattern = new Regex( @"(?:\d*)(?<ss>(\(?(\d{2,4})[-|\.])*(\d{1,4})(\)?))" );

            MatchCollection mc = pattern.Matches( name );
            string num = null;
            if (mc.Count > 0) {
                GroupCollection groups = mc[0].Groups;
                num = groups["ss"].Value;
            }
            if (num != null) {
                int ii = name.IndexOf( num, System.StringComparison.Ordinal );
                var charEnd = name.Length;
                return name.Substring( num.Length + ii );
            }
            return name;
        }

        public string removeBeginNumber(string number) {
            number = "QQ-" + number;
            int len = number.Length;
            int stopNum = len;
            for (int i = len - 1; i > 0; i--) {
                char c = number[i];
                if (( c > 47 ) && ( c < 58 ) || c == '-') {
                    stopNum--;
                } else {
                    break;
                }
            }
            return number.Substring( stopNum + 1 );
        }

        // Локальные сметы. Обработка
        private string WorkWithExcelLs(Worksheet sheet, string selectedItem) {
            // Вытащим из названия стройки номер сметы и одновременно удалим этот номер из названия
            // Локальная смета - ячейка D12
            string numberCellName = "D12";
            string numberSmeta = getNumberSmeta( sheet.Range[numberCellName].Text );
            sheet.Range[numberCellName].Value2 = removeNumberFromNameSmeta( sheet.Range[numberCellName].Text );

            // Найдем заголовок сметы и добавим к нему номер
            Range rangeWork = sheet.Range["A1", "Q14"].Find( @"ЛОКАЛЬНЫЙ СМЕТНЫЙ РАСЧЕТ №" );
            if (rangeWork.Value2 != null) {
                string number = rangeWork.Value2;
                rangeWork.Value2 = number.Remove( number.IndexOf( '№' ) + 1 ) + " " + numberSmeta;
            }
            // Это уровень цен ===============================================
            rangeWork = sheet.Range["A18", "Q21"].Find( "Составлен(а) в текущих (прогнозных) ценах по состоянию на" );
            if (rangeWork != null) {
                int qqq = @"Составлен(а) в текущих (прогнозных) ценах по состоянию на".Length;
                if (qqq >= 0) {
                    string price = @"Составлен(а) в текущих (прогнозных) ценах по состоянию на ";
                    rangeWork.Value2 = price + QuarterFromDate( MainFormAsm.iniSet.CbPriceLevelO );
                    rangeWork.HorizontalAlignment = HorizontalAlignment.Right;
                }
            }
            // Поработаем с подписями
            var firstName = "";
            var secondName = "";
            var stroka1 = 0;
            var stroka2 = 0;
            // Смета нового образца
            // Вначале все очистим от старых
            var range10 = sheet.Cells.Find( @"Составил" );
            if (range10 != null) {
                stroka1 = range10.Row;
                var s1 = range10.Value2 as string;
                if (s1 != null) {
                    firstName = ReturnFio( s1 );
                } // первое имя
            }
            var range20 = sheet.Cells.Find( @"Проверил" );
            if (range20 != null) {
                stroka2 = range20.Row;
                var s2 = range20.Value2 as string;
                if (s2 != null) {
                    secondName = ReturnFio( s2 );
                } // второе имя
            }
            // Очищаем и развоплощаем объединенные ячейки с подписями
            //if (stroka1 != 0 && stroka2 != 0) {
            //    range20 = sheet.Range[sheet.Cells[stroka1, "A"], sheet.Cells[stroka2 + 1, "A"]];
            //    range20.Value2 = "";
            //    range20.UnMerge();
            //    sheet.Range[sheet.Cells[stroka1, "B"], sheet.Cells[stroka2, "Q"]].WrapText = false;
            //    sheet.Cells[stroka1, "E"] = @"Составил :";
            //    sheet.Cells[stroka1, "I"] = firstName;
            //    sheet.Cells[stroka2, "E"] = @"Проверил :";
            //    sheet.Cells[stroka2, "I"] = secondName;
            //} else if (stroka1 == 0 && stroka2 != 0) {
            //    range20 = sheet.Range[sheet.Cells[stroka2, "A"], sheet.Cells[stroka2 + 1, "A"]];
            //    range20.Value2 = "";
            //    range20.UnMerge();
            //    sheet.Range[sheet.Cells[stroka2, "B"], sheet.Cells[stroka2, "Q"]].WrapText = false;
            //    sheet.Cells[stroka2, "E"] = @"Проверил :";
            //    sheet.Cells[stroka2, "I"] = secondName;
            //} else if (stroka1 != 0 && stroka2 == 0) {
            //    range20 = sheet.Range[sheet.Cells[stroka1, "A"], sheet.Cells[stroka1 + 1, "A"]];
            //    range20.Value2 = "";
            //    range20.UnMerge();
            //    sheet.Range[sheet.Cells[stroka1, "B"], sheet.Cells[stroka1, "Q"]].WrapText = false;
            //    sheet.Cells[stroka1, "E"] = @"Составил :";
            //    sheet.Cells[stroka1, "I"] = firstName;
            //}
            //}
            // Вставим подписи в ЛС если нужно
            // Подписи в конце страницы =======================================================
            if (MainFormAsm.iniSet.CbInsertSignLE) {
                // вставим надписи и ФИО
                if (!firstName.Equals( "" ) && stroka1 != 0) {
                    InsertImage( ref sheet, stroka1, 6, firstName );
                }
                if (!secondName.Equals( "" ) && stroka2 != 0) {
                    InsertImage( ref sheet, stroka2, 6, secondName );
                }
            }

            // Подписать страницу
            //sheet.Name = @"ЛС " + numberSmeta;
            sheet.Name = @"ЛС " + removeBeginNumber( numberSmeta );
            // Уберем все лишнее сверху ===============================================
            var range5 = sheet.Range["A1", "Q5"];
            range5.ClearContents();
            //range5.Delete();
            // Это название стройки ===============================================
            rangeWork = sheet.Range["C6", "Q6"];
            rangeWork.MergeCells = true;
            rangeWork.WrapText = true;
            rangeWork.Value2 = MainFormAsm.iniSet.TbNameObject;
            rangeWork.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            rangeWork.Font.Size = 12;
            rangeWork.NumberFormatLocal = "Основной";
            SetRowHeigths( ref sheet, ref rangeWork );
            return numberSmeta;
        }

        private string getNumberSmeta(string text) {
            return text.Substring( 0, text.IndexOf( " " ) );
        }

        private string removeNumberFromNameSmeta(string text) {
            return text.Substring( text.IndexOf( " " ) + 1 );
        }

        private string ReturnFio(string s) {
            var temp = "___________________________";
            return s.Remove( 0, s.IndexOf( temp ) + 27 );
        }

        // Объектные сметы. Обработка
        private string WorkWithExcelOs(Worksheet sheet, string selectedItem) {
            // Затрем все лишнее сверху ===============================================
            var rangeWork = sheet.Range["A1", "Q10"].Find( @"Форма № 3" );
            if (rangeWork != null) {
                rangeWork.Clear();
            }

            // Вытащим из названия стройки номер сметы и одновременно удалим этот номер из названия
            // Локальная смета - ячейка D8
            string numberCellName = "D8";
            string numberSmeta = getNumberSmeta( sheet.Range[numberCellName].Text );
            sheet.Range[numberCellName].Value2 = removeNumberFromNameSmeta( sheet.Range[numberCellName].Text );

            // Это номер сметы ===============================================
            rangeWork = sheet.Cells.Find( @"ОБЪЕКТНЫЙ СМЕТНЫЙ РАСЧЕТ №" );
            if (rangeWork.Value2 != null) {
                string number = rangeWork.Value2;
                rangeWork.Value2 = number + " " + numberSmeta;
            }
            // Капремонт
            if (MainFormAsm.iniSet.CbRebuild) {
                rangeWork = sheet.Cells.Find( @"строительных работ" );
                if (rangeWork != null)
                    rangeWork.Value2 = @"ремонтно-строительных работ";
                rangeWork = sheet.Cells.Find( @"монтажных работ" );
                if (rangeWork != null)
                    rangeWork.Value2 = @"ремонтно-монтажных работ";
                rangeWork = sheet.Cells.Find( @"мебели, инвентаря" );
                if (rangeWork != null)
                    rangeWork.Value2 = @"комплектующих и запасных частей";
                rangeWork = sheet.Cells.Find( @"на строительство" );
                if (rangeWork != null)
                    rangeWork.Value2 = @"";
            }
            // Это уровень цен ===============================================
            // Это уровень цен ===============================================
            rangeWork = sheet.Range["A10", "Q14"].Find( "Составлен(а) в ценах по состоянию на" );
            if (rangeWork != null) {
                int qqq = @"Составлен(а) в ценах по состоянию на".Length;
                if (qqq >= 0) {
                    string price = @"Составлен(а) в ценах по состоянию на ";
                    rangeWork.Value2 = price + QuarterFromDate( MainFormAsm.iniSet.CbPriceLevelO );
                    rangeWork.HorizontalAlignment = HorizontalAlignment.Right;
                }
            }

            //rangeWork = sheet.Cells.Find(@"Составлен  в ценах по состоянию на");
            //string price = rangeWork.Text;
            //int qqq = @"Составлен  в ценах по состоянию на".Length;
            //if (qqq >= 0)
            //{
            //    price = @"Составлена в ценах по состоянию на ";
            //    rangeWork.Value2 = price + QuarterFromDate(MainFormAsm.iniSet.CbPriceLevelO);
            //    rangeWork.HorizontalAlignment = HorizontalAlignment.Right;
            //}
            RewriteFirstStringTable( sheet );
            // Подписи в конце страницы =======================================================
            if (MainFormAsm.iniSet.CbInsertSignOE) {
                var rowEnd = FindLastRow( sheet );
                var rowGip = rowEnd + 4;
                var rowBoss = rowEnd + 7;
                var rowMadeIn = rowEnd + 10;
                // вставим надписи и ФИО
                rangeWork = sheet.Cells[rowGip, "C"];
                rangeWork.Value2 = @"Главный инженер проекта";
                rangeWork.HorizontalAlignment = XlHAlign.xlHAlignRight;
                rangeWork = sheet.Cells[rowBoss, "C"];
                rangeWork.Value2 = @"Начальник отдела";
                rangeWork.HorizontalAlignment = XlHAlign.xlHAlignRight;
                rangeWork = sheet.Cells[rowMadeIn, "C"];
                rangeWork.Value2 = @"Составил";
                rangeWork.HorizontalAlignment = XlHAlign.xlHAlignRight;

                rangeWork = sheet.Cells[rowGip, "F"];
                rangeWork.Value2 = "______________________" + MainFormAsm.iniSet.CbGipText;
                rangeWork.HorizontalAlignment = XlHAlign.xlHAlignRight;
                rangeWork = sheet.Cells[rowBoss, "F"];
                rangeWork.Value2 = "______________________" + MainFormAsm.iniSet.TbHeadDepartment;
                rangeWork.HorizontalAlignment = XlHAlign.xlHAlignRight;
                rangeWork = sheet.Cells[rowMadeIn, "F"];
                rangeWork.Value2 = "______________________" + MainFormAsm.iniSet.CbMadeInText;
                rangeWork.HorizontalAlignment = XlHAlign.xlHAlignRight;
                // Вставим картинки
                if (!MainFormAsm.iniSet.CbGip.Equals( "" )) {
                    InsertImage( ref sheet, rowGip, 5, MainFormAsm.iniSet.CbGipText );
                }
                if (!MainFormAsm.iniSet.TbHeadDepartment.Equals( "" )) {
                    InsertImage( ref sheet, rowBoss, 5, MainFormAsm.iniSet.TbHeadDepartment );
                }
                if (!MainFormAsm.iniSet.CbMadeIn.Equals( "" )) {
                    InsertImage( ref sheet, rowMadeIn, 5, MainFormAsm.iniSet.CbMadeInText );
                }
            }
            sheet.Name = @"ОС " + removeBeginNumber( numberSmeta );
            // Это название стройки ===============================================
            rangeWork = sheet.Range["C2", "I2"];
            rangeWork.MergeCells = true;
            rangeWork.WrapText = true;
            rangeWork.Value2 = MainFormAsm.iniSet.TbNameObject;
            rangeWork.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            rangeWork.Font.Size = 12;
            rangeWork.NumberFormatLocal = "Основной";
            SetRowHeigths( ref sheet, ref rangeWork );
            return sheet.Name;
        }

        // Переделка начальных строк таблицы
        private void RewriteFirstStringTable(_Worksheet sheet) {
            Range col = sheet.Columns[2];
            col.ColumnWidth = "14";
            Range find = sheet.Cells.Find( @"Локальные сметные расчеты" );
            if (find == null) {
                return;
            }
            int end = sheet.Cells.Find( "Итого \"Локальные сметные расчеты\"" ).Row;
            int y = find.Row + 1;
            int x1 = find.Column + 1;
            int x2 = x1 + 1;
            for (int i = y; i < end; i++) {
                Range r1 = sheet.Cells[i, x1];
                Range r2 = sheet.Cells[i, x2];
                string s1 = r1.Value2;
                string s2 = r2.Value2;
                if (s1 != null && s2 != null) {
                    int x10 = s2.IndexOf( '(' );
                    int x11 = s2.IndexOf( ')' );
                    if (x10 >= 0 && x11 > 0) {
                        s1 = s2.Substring( s2.IndexOf( "(", System.StringComparison.Ordinal ) + 1,
                            s2.IndexOf( ")", System.StringComparison.Ordinal ) - 1 );
                        s2 = s2.Substring( s2.IndexOf( ")", System.StringComparison.Ordinal ) + 1 );
                        r1.Value2 = s1;
                        r2.Value2 = s2;
                    }
                }
            }
        }

        // Ресурсы. Обработка
        private string WorkWithExcelR(Worksheet sheet) {
            // Это наименование работ и т.д. ===============================================
            Range rangeWork = sheet.Range["C8", "C8"];
            string nameWorks = RenameName( rangeWork.Value2 );
            if (rangeWork.Value2 != null) {
                string sss = rangeWork.Value2.ToString();
                rangeWork.Value2 = RemoveBeginPos( sss );
            }
            // Это номер сметы ===============================================
            rangeWork = sheet.Cells.Find( @"ВЕДОМОСТЬ РЕСУРСОВ" );
            if (rangeWork.Value2 != null) {
                string number = rangeWork.Value2;
                rangeWork.Value2 = number + " " + nameWorks;
            }
            // Это уровень цен ===============================================
            rangeWork = sheet.Cells.Find( @"по состоянию на" );
            string price = rangeWork.Value2;
            rangeWork.Value2 = price + " " + QuarterFromDate( MainFormAsm.iniSet.CbPriceLevelL );
            //            rangeWork.Value2 = price + " " + cbPriceLevel.Text;
            rangeWork.HorizontalAlignment = HorizontalAlignment.Right;
            // Имя файла
            sheet.Name = @"Р" + nameWorks;
            // Это название стройки ===============================================
            rangeWork = sheet.Range["B2", "H2"];
            rangeWork.MergeCells = true;
            rangeWork.WrapText = true;
            rangeWork.Value2 = MainFormAsm.iniSet.TbNameObject;
            rangeWork.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            rangeWork.Font.Size = 12;
            rangeWork.NumberFormatLocal = "Основной";
            SetRowHeigths( ref sheet, ref rangeWork );
            return nameWorks;
        }

        private string WorkWithExcelVR(_Worksheet sheet) {
            // Это наименование работ и т.д. ===============================================
            var rangeWork = sheet.Range["C7", "C7"];
            string nameWorks = RenameName( rangeWork.Value2 );
            // Убираем знак "№" из заголовка ===============================================
            rangeWork = sheet.Cells.Find( @"ВЕДОМОСТЬ ОБЪЕМОВ РАБОТ №" );
            if (rangeWork.Value2 != null) {
                string number = rangeWork.Value2;
                rangeWork.Value2 = number.Remove( number.IndexOf( "№", System.StringComparison.Ordinal ) );
            }
            // Уберем все лишнее сверху ===============================================
            var range5 = sheet.Range["A1", "E4"];
            range5.ClearContents();
            // Имя страницы
            sheet.Name = @"ВР" + nameWorks;
            return nameWorks;
        }

        private void SetRowHeigths(ref Worksheet ws, ref Range src) {
            Range test = ws.Cells[900, 100];
            int aa = src.EntireColumn.Count;
            double colWidth = 0;
            for (int i = 1; i <= aa; i++) {
                Range r = src.EntireColumn[i];
                colWidth = colWidth + r.ColumnWidth;
            }
            test.ColumnWidth = colWidth;
            test.Font.Size = 12;
            test.Value2 = src.Value2;
            test.WrapText = true;
            test.Rows.AutoFit();
            double h = test.RowHeight;
            h = Math.Ceiling( h / 10 ) * 10;
            src.RowHeight = h;
            test.Delete();
        }

        private float FloatTopPixelsCalculation(Range range) {
            Range r1 = range.Worksheet.Cells[range.Row + 1, range.Column];
            float floatTop1 = 0;
            for (var rNumber = 2; rNumber < r1.Row; rNumber++) {
                var cellHeight = Convert.ToSingle( r1.Worksheet.Cells[rNumber, r1.Column].RowHeight );
                floatTop1 = floatTop1 + cellHeight;
            }
            float floatTop = 0;
            for (var rNumber = 2; rNumber < range.Row; rNumber++) {
                var cellHeight = Convert.ToSingle( range.Worksheet.Cells[rNumber, range.Column].RowHeight );
                floatTop = floatTop + cellHeight;
            }
            return ( floatTop + floatTop1 ) / 2;
        }

        private float FloatLeftPixelsCalculation(Range range) {
            float floatLeft = 0;
            for (var columnNumber = 1; columnNumber < range.Columns.Column; columnNumber++) {
                var cellWidth = Convert.ToSingle( range.Worksheet.Cells[range.Row, columnNumber].Width );
                floatLeft = floatLeft + cellWidth + 1;
            }
            return floatLeft;
        }
    }
}
