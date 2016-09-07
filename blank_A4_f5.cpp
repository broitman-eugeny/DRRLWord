//Функции класса бланка формата А4 с основной надписью по форме 5


#include <afx.h>
#include "afxdisp.h"//Необходима для работы AfxOleInit()
#include "msword9.h"//Заголовочный файл, полученный с помощью ClassWizard Visual Studio
#include "DMRWord.h"//Заголовочный файл с описаниями классов MS Word



cBlank_A4_f5::cBlank_A4_f5()//пустой конструктор
{
}





//Инициализирует систему OLE. Обеспечивает поддержку элементов управления OLE.
//Запускает сервер автоматизации "Word.Application".
//Отображает основную надпись и рамку в нижнем колонтитуле на текущем листе открытого документа Word
//Исходные данные для заполнения основной надписи берет по ссылке на объект типа cData_interval_DMR
//Описание класса cData_interval_DMR хранится в в файле DMR.h
//В случае успешного построения бланка возвращает TRUE, в случае неудачи или прерывания
//пользователем - FALSE.
BOOL	cBlank_A4_f5::Draw_blank(cData_interval_DMR *cDiD)
{




   //Это заставит проинициализироваться систему OLE. Если этого не сделать, то вызов CreateDispatch не сработает.
   /*Должна вызываться приложением. В DLL выдает ошибку
   if(!AfxOleInit()) // Your addition starts here
   {
   	AfxMessageBox((LPCTSTR)"Could not initialize COM dll",(UINT)MB_OK,(UINT)0);
      return FALSE;
   }               // End of your addition
   */

   AfxEnableControlContainer();//Call this function in application object's InitInstance function to enable support for
      									//containment of OLE controls.

   if(!app.CreateDispatch("Word.Application")) //запустить сервер
   {
   	AfxMessageBox((LPCTSTR)"Ошибка при старте Wordа!",(UINT)MB_OK,(UINT)0);
      return FALSE;
   }


   app.SetVisible(TRUE); //и сделать Word видимым


   COleVariant  covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
   //наша коллекция документов
   oDocs = app.GetDocuments();
   //добавить к ней новый документ
   //Внимание! Если у вас Word 97 - то строчка будет такая:
   //oDocs.Add(covOptional,covOptional);				//97
   oDocs.Add(covOptional,covOptional,covOptional,covOptional);	//2000
   //и получить его как экзепляр коллекции с номером 1
   Word_blank = oDocs.Item(COleVariant(long(1)));
   //активизировать документ
   Word_blank.Activate();

   //Установка параметров страницы
   CPageSetup oPageSetup;
   oPageSetup=this->Word_blank.GetPageSetup();
   oPageSetup.put_DifferentFirstPageHeaderFooter(long(True));//Разные колонтитулы первого и остальных листов
   oPageSetup.put_LeftMargin(20.*MM2PH);//Расстояние от левого края листа
   oPageSetup.put_RightMargin(5.*MM2PH);//Расстояние от правого края листа
   oPageSetup.put_TopMargin(10.*MM2PV);//Расстояние от верхнего края листа
   oPageSetup.put_BottomMargin(5.*MM2PV);//Расстояние от нижнего края листа
   oPageSetup.put_HeaderDistance(0.*MM2PV);//Расстояние от верхнего края листа до верхнего колонтитула
   oPageSetup.put_FooterDistance(5.*MM2PV);//Расстояние от нижнего края листа до нижнего колонтитула
   Word_blank.SetPageSetup(oPageSetup);

   //Установка вида 100%
   CWindow0 ActiveWindow;
   ActiveWindow=this->Word_blank.GetActiveWindow();
   CPane ActivePane;
   ActivePane=ActiveWindow.get_ActivePane();
   CView0 View;
   View=ActivePane.get_View();
   CZoom Zoom;
   Zoom=View.get_Zoom();
   Zoom.put_Percentage(100);

   //Вход в нижний колонтитул (footer)
   View.put_SeekView(wdSeekCurrentPageFooter);

   //Создаем рамку надписи
   Selection oSel;
   oSel = app.GetSelection();
   CHeaderFooter oHeaderFooter;
   oHeaderFooter=oSel.GetHeaderFooter();
   CShapes oShapes;
   oShapes=oHeaderFooter.get_Shapes();
   CShape oShape;
   Range oRan;
   oShape=oShapes.AddShape(msoShapeRectangle,5.*MM2PH,145.6*MM2PV,15.*MM2PH,158.4*MM2PV,covOptional);
   CFillFormat oFillFormat;
   CLineFormat oLineFormat;
   oFillFormat=oShape.get_Fill();
   oLineFormat=oShape.get_Line();
   oFillFormat.put_Visible(long (FALSE));
   oLineFormat.put_Visible(long (FALSE));
   CTextFrame oTextFrame;
   oTextFrame=oShape.get_TextFrame();
   oTextFrame.put_MarginBottom(0.);
   oTextFrame.put_MarginLeft(0);
   oTextFrame.put_MarginRight(0.);
   oTextFrame.put_MarginTop(0.);

   //Создаем таблицу бокового штампа
   Tables oTables;
   Table oTable;
   oRan = oTextFrame.get_TextRange();
   oTables = this->Word_blank.GetTables();

   //добавить таблицу в коллекцию
   oTable = oTables.Add(oRan,7,5,COleVariant(short(wdWord9TableBehavior)),COleVariant(short(wdAutoFitFixed)));
   //Установка направления текста в таблице
   oTable.Select();
   oSel.SetOrientation(wdTextOrientationUpward);

   //Установка шрифта в таблице
   _Font oFont;
   oFont=oSel.GetFont();
   oFont.SetSize(10);
   oFont.SetName("Arial");

   //Устанавливаем минимальные поля ячеек таблицы
	oTable.SetTopPadding(0.);
   oTable.SetBottomPadding(0.);
   oTable.SetLeftPadding(0.);
   oTable.SetRightPadding(0.);

   //Устанавливаем минимальное расстояние между ячейками таблицы
   oTable.SetSpacing(0.);
   //Устанавливаем автоподгонку размеров ячеек под содержимое
   oTable.SetAllowAutoFit(BOOL(true));

   //Устанавливаем вертикальное выравнивание в ячейках таблицы
   Cells oCells;
   oTable.Select();
   oCells=oSel.GetCells();
   oCells.SetVerticalAlignment(wdCellAlignVerticalCenter);

   //Устанавливаем высоту строк
   Rows oRows;
   oRows=oTable.GetRows();
   Row oRow;
   oRow=oRows.Item(1);
   oRow.SetHeight(10.*MM2PV);
   oRow=oRows.Item(2);
   oRow.SetHeight(15.*MM2PV);
   oRow=oRows.Item(3);
   oRow.SetHeight(20.*MM2PV);
   oRow=oRows.Item(4);
   oRow.SetHeight(20.*MM2PV);
   oRow=oRows.Item(5);
   oRow.SetHeight(25.*MM2PV);
   oRow=oRows.Item(6);
   oRow.SetHeight(35.*MM2PV);
   oRow=oRows.Item(7);
   oRow.SetHeight(25.*MM2PV);

   //Устанавливаем ширину столбцов
   Columns oColumns;
   oColumns=oTable.GetColumns();
   Column oColumn;
   oColumn=oColumns.Item(1);
   oColumn.SetPreferredWidth(3.*MM2PH);
   oColumn=oColumns.Item(2);
   oColumn.SetPreferredWidth(2.*MM2PH);
   oColumn=oColumns.Item(3);
   oColumn.SetPreferredWidth(3.*MM2PH);
   oColumn=oColumns.Item(4);
   oColumn.SetPreferredWidth(2.*MM2PH);
   oColumn=oColumns.Item(5);
   oColumn.SetPreferredWidth(5.*MM2PH);

   //Объединяем ячейки
   Cell oCell;
   oCell = oTable.Cell(1,1);
   oCell.Select();
   oSel.MoveDown(COleVariant(short(wdLine)),COleVariant(long(3)),COleVariant(short(wdExtend)));
   oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdExtend)));
   oCells=oSel.GetCells();
   oCells.Merge();
   oCell = oTable.Cell(1,2);
   oCell.Select();
   oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdExtend)));
   oCells=oSel.GetCells();
   oCells.Merge();
   oCell = oTable.Cell(2,2);
   oCell.Select();
   oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdExtend)));
   oCells=oSel.GetCells();
   oCells.Merge();
   oCell = oTable.Cell(3,2);
   oCell.Select();
   oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdExtend)));
   oCells=oSel.GetCells();
   oCells.Merge();
   oCell = oTable.Cell(4,2);
   oCell.Select();
   oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdExtend)));
   oCells=oSel.GetCells();
   oCells.Merge();
   oCell = oTable.Cell(5,2);
   oCell.Select();
   oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdExtend)));
   oCells=oSel.GetCells();
   oCells.Merge();
   oCell = oTable.Cell(6,2);
   oCell.Select();
   oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdExtend)));
   oCells=oSel.GetCells();
   oCells.Merge();
   oCell = oTable.Cell(7,2);
   oCell.Select();
   oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdExtend)));
   oCells=oSel.GetCells();
   oCells.Merge();
   oCell = oTable.Cell(5,3);
   oCell.Select();
   oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdExtend)));
   oCells=oSel.GetCells();
   oCells.Merge();
   oCell = oTable.Cell(6,3);
   oCell.Select();
   oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdExtend)));
   oCells=oSel.GetCells();
   oCells.Merge();
   oCell = oTable.Cell(7,3);
   oCell.Select();
   oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdExtend)));
   oCells=oSel.GetCells();
   oCells.Merge();

   //Устанавливаем толщину границ ячеек таблицы
   Borders oBorders;
   Border oBorder;
   oBorders=oTable.GetBorders();
   oBorder=oBorders.Item(wdBorderLeft);
   oBorder.SetLineWidth(wdLineWidth150pt);
   oBorder=oBorders.Item(wdBorderTop);
   oBorder.SetLineWidth(wdLineWidth150pt);
   oBorder=oBorders.Item(wdBorderBottom);
   oBorder.SetLineWidth(wdLineWidth150pt);
   oBorder=oBorders.Item(wdBorderRight);
   oBorder.SetLineWidth(wdLineWidth150pt);
   oBorder=oBorders.Item(wdBorderHorizontal);
   oBorder.SetLineWidth(wdLineWidth150pt);
   oBorder=oBorders.Item(wdBorderVertical);
   oBorder.SetLineWidth(wdLineWidth150pt);
   oCell = oTable.Cell(5,1);
   oCell.Select();
   oSel.MoveDown(COleVariant(short(wdLine)),COleVariant(long(2)),COleVariant(short(wdExtend)));
   oBorders=oSel.GetBorders();
   oBorder=oBorders.Item(wdBorderLeft);
   oBorder.SetVisible(FALSE);
   oBorder=oBorders.Item(wdBorderHorizontal);
   oBorder.SetVisible(FALSE);
   oBorder=oBorders.Item(wdBorderBottom);
   oBorder.SetVisible(FALSE);
   oCell = oTable.Cell(1,2);
   oCell.Select();
   oSel.MoveDown(COleVariant(short(wdLine)),COleVariant(long(3)),COleVariant(short(wdExtend)));
   oBorders=oSel.GetBorders();
   oBorder=oBorders.Item(wdBorderRight);
   oBorder.SetLineWidth(wdLineWidth075pt);

   //Проставляем надписи в графах
   //Графа "Инв. № подл."
   oCell = oTable.Cell(7,2);
   oRan = oCell.GetRange();
   oRan.SetText("Инв. № подл.");
   oCell.Select();
   oFont.SetSize(8);
   oSel.BoldRun();
   Paragraphs oPars;
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);

   //Графа 20
   oCell = oTable.Cell(7,3);
   oRan = oCell.GetRange();
   oRan.SetText(cDiD->sN_podl);
   oCell.Select();
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);

   //Графа "Подп. и дата"
   oCell = oTable.Cell(6,2);
   oRan = oCell.GetRange();
   oRan.SetText("Подп. и дата");
   oCell.Select();
   oFont.SetSize(8);
   oSel.BoldRun();
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);

   //Графа "Взам. инв. №"
   oCell = oTable.Cell(5,2);
   oRan = oCell.GetRange();
   oRan.SetText("Взам. инв. №");
   oCell.Select();
   oFont.SetSize(8);
   oSel.BoldRun();
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);

   //Графа 22
   oCell = oTable.Cell(5,3);
   oRan = oCell.GetRange();
   oRan.SetText(cDiD->sN_star_podl);
   oCell.Select();
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);

   //Графа "Согласовано"
   oCell = oTable.Cell(1,1);
   oRan = oCell.GetRange();
   oRan.SetText("Согласовано");
   oCell.Select();
   oFont.SetSize(8);
   oSel.BoldRun();
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_LEFT);
   if(cDiD->sFam_gip!="")
   {
   	//Графа 10 в боковом штампе
   	oCell = oTable.Cell(4,2);
   	oRan = oCell.GetRange();
   	oRan.SetText("ГИП");
   	oCell.Select();
   	oPars=oRan.GetParagraphs();
   	oPars.SetAlignment(AL_LEFT);
      //Графа 11 в боковом штампе
   	oCell = oTable.Cell(3,2);
   	oRan = oCell.GetRange();
   	oRan.SetText(cDiD->sFam_gip);
   	oCell.Select();
   	oPars=oRan.GetParagraphs();
   	oPars.SetAlignment(AL_LEFT);
   }


   //Создаем таблицу основного штампа
   View.put_SeekView(wdSeekCurrentPageFooter);
   oRan = oSel.GetRange();
   oTables = this->Word_blank.GetTables();

   //добавить таблицу в коллекцию
   oTable = oTables.Add(oRan,8,10,COleVariant(short(wdWord9TableBehavior)),COleVariant(short(wdAutoFitFixed)));
   //Установка шрифта в таблице
   oTable.Select();
   oFont.SetSize(10);
   oFont.SetName("Arial");

   //Устанавливаем минимальные поля ячеек таблицы
	oTable.SetTopPadding(0.);
   oTable.SetBottomPadding(0.);
   oTable.SetLeftPadding(0.);
   oTable.SetRightPadding(0.);

   //Устанавливаем минимальное расстояние между ячейками таблицы
   oTable.SetSpacing(0.);
   //Устанавливаем автоподгонку размеров ячеек под содержимое
   oTable.SetAllowAutoFit(BOOL(true));

   //Устанавливаем вертикальное выравнивание в ячейках таблицы
   oTable.Select();
   oCells=oSel.GetCells();
   oCells.SetVerticalAlignment(wdCellAlignVerticalCenter);

   //Устанавливаем высоту строк
   oRows=oTable.GetRows();
   oRows.SetHeight(5.*MM2PV*kMM2PV);

   //Устанавливаем ширину столбцов
   oColumns=oTable.GetColumns();
   oColumn=oColumns.Item(1);
   oColumn.SetPreferredWidth(10.*MM2PH);
   oColumn=oColumns.Item(2);
   oColumn.SetPreferredWidth(10.*MM2PH);
   oColumn=oColumns.Item(3);
   oColumn.SetPreferredWidth(10.*MM2PH);
   oColumn=oColumns.Item(4);
   oColumn.SetPreferredWidth(10.*MM2PH);
   oColumn=oColumns.Item(5);
   oColumn.SetPreferredWidth(15.*MM2PH);
   oColumn=oColumns.Item(6);
   oColumn.SetPreferredWidth(10.*MM2PH);
   oColumn=oColumns.Item(7);
   oColumn.SetPreferredWidth(70.*MM2PH);
   oColumn=oColumns.Item(8);
   oColumn.SetPreferredWidth(15.*MM2PH);
   oColumn=oColumns.Item(9);
   oColumn.SetPreferredWidth(15.*MM2PH);
   oColumn=oColumns.Item(10);
   oColumn.SetPreferredWidth(20.*MM2PH);

   //Объединяем ячейки
   oCell = oTable.Cell(4,1);
   oCell.Select();
   oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdExtend)));
   oCells=oSel.GetCells();
   oCells.Merge();
   oCell = oTable.Cell(5,1);
   oCell.Select();
   oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdExtend)));
   oCells=oSel.GetCells();
   oCells.Merge();
   oCell = oTable.Cell(6,1);
   oCell.Select();
   oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdExtend)));
   oCells=oSel.GetCells();
   oCells.Merge();
   oCell = oTable.Cell(7,1);
   oCell.Select();
   oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdExtend)));
   oCells=oSel.GetCells();
   oCells.Merge();
   oCell = oTable.Cell(8,1);
   oCell.Select();
   oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdExtend)));
   oCells=oSel.GetCells();
   oCells.Merge();
   oCell = oTable.Cell(4,2);
   oCell.Select();
   oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdExtend)));
   oCells=oSel.GetCells();
   oCells.Merge();
   oCell = oTable.Cell(5,2);
   oCell.Select();
   oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdExtend)));
   oCells=oSel.GetCells();
   oCells.Merge();
   oCell = oTable.Cell(6,2);
   oCell.Select();
   oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdExtend)));
   oCells=oSel.GetCells();
   oCells.Merge();
   oCell = oTable.Cell(7,2);
   oCell.Select();
   oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdExtend)));
   oCells=oSel.GetCells();
   oCells.Merge();
   oCell = oTable.Cell(8,2);
   oCell.Select();
   oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdExtend)));
   oCells=oSel.GetCells();
   oCells.Merge();
   oCell = oTable.Cell(1,7);
   oCell.Select();
   oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(3)),COleVariant(short(wdExtend)));
   oSel.MoveDown(COleVariant(short(wdLine)),COleVariant(long(2)),COleVariant(short(wdExtend)));
   oCells=oSel.GetCells();
   oCells.Merge();
   oCell = oTable.Cell(4,5);
   oCell.Select();
   oSel.MoveDown(COleVariant(short(wdLine)),COleVariant(long(4)),COleVariant(short(wdExtend)));
   oCells=oSel.GetCells();
   oCells.Merge();
   oCell = oTable.Cell(6,6);
   oCell.Select();
   oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdExtend)));
   oSel.MoveDown(COleVariant(short(wdLine)),COleVariant(long(2)),COleVariant(short(wdExtend)));
   oCells=oSel.GetCells();
   oCells.Merge();

   //Устанавливаем толщину границ ячеек таблицы
   oBorders=oTable.GetBorders();
   oBorder=oBorders.Item(wdBorderLeft);
   oBorder.SetLineWidth(wdLineWidth150pt);
   oBorder=oBorders.Item(wdBorderTop);
   oBorder.SetLineWidth(wdLineWidth150pt);
   oBorder=oBorders.Item(wdBorderBottom);
   oBorder.SetLineWidth(wdLineWidth150pt);
   oBorder=oBorders.Item(wdBorderRight);
   oBorder.SetLineWidth(wdLineWidth150pt);
   oBorder=oBorders.Item(wdBorderHorizontal);
   oBorder.SetLineWidth(wdLineWidth150pt);
   oBorder=oBorders.Item(wdBorderVertical);
   oBorder.SetLineWidth(wdLineWidth150pt);
   oCell = oTable.Cell(1,1);
   oCell.Select();
   oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(5)),COleVariant(short(wdExtend)));
   oSel.MoveDown(COleVariant(short(wdLine)),COleVariant(long(1)),COleVariant(short(wdExtend)));
   oBorders=oSel.GetBorders();
   oBorder=oBorders.Item(wdBorderHorizontal);
   oBorder.SetLineWidth(wdLineWidth075pt);
   oCell = oTable.Cell(4,1);
   oCell.Select();
   oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(3)),COleVariant(short(wdExtend)));
   oSel.MoveDown(COleVariant(short(wdLine)),COleVariant(long(4)),COleVariant(short(wdExtend)));
   oBorders=oSel.GetBorders();
   oBorder=oBorders.Item(wdBorderHorizontal);
   oBorder.SetLineWidth(wdLineWidth075pt);

   //Проставляем надписи в графах
   //Графа "Изм."
   oCell = oTable.Cell(3,1);
   oRan = oCell.GetRange();
   oRan.SetText("Изм.");
   oCell.Select();
   oFont.SetSize(7);
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);

   //Графа "Кол. уч."
   oCell = oTable.Cell(3,2);
   oRan = oCell.GetRange();
   oRan.SetText("Кол. уч.");
   oCell.Select();
   oFont.SetSize(7);
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);

   //Графа "Лист" в таблице изменений
   oCell = oTable.Cell(3,3);
   oRan = oCell.GetRange();
   oRan.SetText("Лист");
   oCell.Select();
   oFont.SetSize(7);
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);

   //Графа "№ док."
   oCell = oTable.Cell(3,4);
   oRan = oCell.GetRange();
   oRan.SetText("№ док.");
   oCell.Select();
   oFont.SetSize(7);
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);

   //Графа "Подп." в таблице изменений
   oCell = oTable.Cell(3,5);
   oRan = oCell.GetRange();
   oRan.SetText("Подп.");
   oCell.Select();
   oFont.SetSize(7);
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);

   //Графа "Дата" в таблице изменений
   oCell = oTable.Cell(3,6);
   oRan = oCell.GetRange();
   oRan.SetText("Дата");
   oCell.Select();
   oFont.SetSize(7);
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);
   if(cDiD->sN_izm!="")
   {
   	//Графа 14
   	oCell = oTable.Cell(2,1);
   	oRan = oCell.GetRange();
   	oRan.SetText(cDiD->sN_izm);
   	oCell.Select();
      oFont.SetSize(8);
   	oPars=oRan.GetParagraphs();
   	oPars.SetAlignment(AL_CENTER);
      //Графа 15
   	oCell = oTable.Cell(2,2);
   	oRan = oCell.GetRange();
   	oRan.SetText(cDiD->sKol_uch);
   	oCell.Select();
      oFont.SetSize(8);
   	oPars=oRan.GetParagraphs();
   	oPars.SetAlignment(AL_CENTER);
      //Графа 16
   	oCell = oTable.Cell(2,3);
   	oRan = oCell.GetRange();
   	oRan.SetText(cDiD->sZam_nov_vse);
   	oCell.Select();
      oFont.SetSize(8);
   	oPars=oRan.GetParagraphs();
   	oPars.SetAlignment(AL_CENTER);
      //Графа 17
   	oCell = oTable.Cell(2,4);
   	oRan = oCell.GetRange();
   	oRan.SetText(cDiD->sOb_razr);
   	oCell.Select();
      oFont.SetSize(8);
   	oPars=oRan.GetParagraphs();
   	oPars.SetAlignment(AL_CENTER);
   }

   //Графа "Разраб"
   oCell = oTable.Cell(4,1);
   oRan = oCell.GetRange();
   oRan.SetText("Разраб.");
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_LEFT);

   //Графа 11 фамилия исполнителя
   oCell = oTable.Cell(4,2);
   oRan = oCell.GetRange();
   oRan.SetText(cDiD->sFam_isp);
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_LEFT);
   if(cDiD->sFam_prov!="")
   {
   	//Графа "Проверил"
   	oCell = oTable.Cell(5,1);
   	oRan = oCell.GetRange();
   	oRan.SetText("Проверил");
   	oPars=oRan.GetParagraphs();
   	oPars.SetAlignment(AL_LEFT);
   	//Графа 11 фамилия проверившего
   	oCell = oTable.Cell(5,2);
   	oRan = oCell.GetRange();
   	oRan.SetText(cDiD->sFam_prov);
   	oPars=oRan.GetParagraphs();
   	oPars.SetAlignment(AL_LEFT);
   }

   if(cDiD->sFam_glt!="")
   {
   	//Графа "Гл. техн."
   	oCell = oTable.Cell(6,1);
   	oRan = oCell.GetRange();
   	oRan.SetText("Гл. техн.");
   	oPars=oRan.GetParagraphs();
   	oPars.SetAlignment(AL_LEFT);
   	//Графа 11 фамилия главного технолога
   	oCell = oTable.Cell(6,2);
   	oRan = oCell.GetRange();
   	oRan.SetText(cDiD->sFam_glt);
   	oPars=oRan.GetParagraphs();
   	oPars.SetAlignment(AL_LEFT);
   }

   //Графа "Н. контр."
   oCell = oTable.Cell(7,1);
   oRan = oCell.GetRange();
   oRan.SetText("Н. контр.");
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_LEFT);

   //Графа 11 фамилия нормоконтролера
   oCell = oTable.Cell(7,2);
   oRan = oCell.GetRange();
   oRan.SetText(cDiD->sFam_nk);
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_LEFT);

   if(cDiD->sFam_no!="")
   {
   	//Графа "Нач. отд."
   	oCell = oTable.Cell(8,1);
   	oRan = oCell.GetRange();
   	oRan.SetText("Нач. отд.");
   	oPars=oRan.GetParagraphs();
   	oPars.SetAlignment(AL_LEFT);
   	//Графа 11 фамилия начальника отдела
   	oCell = oTable.Cell(8,2);
   	oRan = oCell.GetRange();
   	oRan.SetText(cDiD->sFam_no);
   	oPars=oRan.GetParagraphs();
   	oPars.SetAlignment(AL_LEFT);
   }

   //Графа 1
   oCell = oTable.Cell(1,7);
   oRan = oCell.GetRange();
   oRan.SetText(cDiD->sObozn_doc);
   oCell.Select();
   oFont.SetSize(14);
   oSel.BoldRun();
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);

   //Графа 5
   oCell = oTable.Cell(4,5);
   oRan = oCell.GetRange();
   oRan.SetText(cDiD->sNaim_doc);
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);

   //Графа "Стадия"
   oCell = oTable.Cell(4,6);
   oRan = oCell.GetRange();
   oRan.SetText("Стадия");
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);

   //Графа "Лист"
   oCell = oTable.Cell(4,7);
   oRan = oCell.GetRange();
   oRan.SetText("Лист");
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);

   //Графа "Листов"
   oCell = oTable.Cell(4,8);
   oRan = oCell.GetRange();
   oRan.SetText("Листов");
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);

   //Графа 6
   oCell = oTable.Cell(5,6);
   oRan = oCell.GetRange();
   oRan.SetText(cDiD->sVid_doc);
   oCell.Select();
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);

   //Графа 7
   CString sN_list_pref="";//префикс нумерации страниц
   long lsn;//стартовый номер страниц

   int index;
   if((index=cDiD->sN_list.Find('-'))!=-1)//Если строка содержит символ '-'
   {
   	sN_list_pref=cDiD->sN_list.Left(index+1);//префикс нумерации страниц
      lsn=atol(&((LPCTSTR(cDiD->sN_list))[index+1]));//стартовый номер страниц
   }
   else lsn=atol(LPCTSTR(cDiD->sN_list));//стартовый номер страниц
   oCell = oTable.Cell(5,7);
   oCell.Select();
   oRan = oSel.GetRange();
   oSel.Collapse(COleVariant(short(wdCollapseStart)));
   oSel.SetText(sN_list_pref);
   oSel.Collapse(COleVariant(short(wdCollapseEnd)));
   CFields oFields;
   oFields=oSel.GetFields();
   oFields.Add(oSel.GetRange(), COleVariant(short(wdFieldPage)), covOptional, covOptional);//Вставка поля "номер страницы"
   CPageNumbers oPN;
   oPN=oHeaderFooter.get_PageNumbers();
   oPN.put_RestartNumberingAtSection(TRUE);
   oPN.put_StartingNumber(lsn);
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);

   //Графа 8
   oCell = oTable.Cell(5,8);
   oRan = oCell.GetRange();
   oRan.SetText(cDiD->sKol_list);
   oCell.Select();
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);

   //Графа 9
   oCell = oTable.Cell(6,6);
   oRan = oCell.GetRange();
   oRan.SetText(cDiD->sNaim_razr);
   oCell.Select();
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);


   //Рисуем прямоугольную рамку
   oShape=oShapes.AddShape(msoShapeRectangle,20.*2.82,5.*2.787,185.*2.834,287.*2.787,covOptional);
   oFillFormat=oShape.get_Fill();
   oLineFormat=oShape.get_Line();
   oFillFormat.put_Visible(long (FALSE));
   oLineFormat.put_Weight(0.5*MM2PH);


   //Добавляем страницу
   View.put_SeekView(wdSeekMainDocument);//Выйти из колонтитула
   COleVariant covBreakType((long)BR_PAGE);
   oSel.InsertBreak(covBreakType);


   //Вход в нижний колонтитул 2-й страницы
   View.put_SeekView(wdSeekCurrentPageFooter);
   oSel.MoveDown(COleVariant(short(wdWindow)),COleVariant(long(1)),COleVariant(short(wdMove)));
   oSel.MoveDown(COleVariant(short(wdLine)),COleVariant(long(2)),COleVariant(short(wdMove)));


   //Создаем рамку надписи для бокового штампа на 2-й странице
   oHeaderFooter=oSel.GetHeaderFooter();
   oShapes=oHeaderFooter.get_Shapes();
   oShape=oShapes.AddShape(msoShapeRectangle,8.*MM2PH,212.4*MM2PV,12.*MM2PH,93.4*MM2PV,covOptional);
	oFillFormat=oShape.get_Fill();
   oLineFormat=oShape.get_Line();
   oFillFormat.put_Visible(long (FALSE));
   oLineFormat.put_Visible(long (FALSE));
   oTextFrame=oShape.get_TextFrame();
   oTextFrame.put_MarginBottom(0.);
   oTextFrame.put_MarginLeft(0);
   oTextFrame.put_MarginRight(0.);
   oTextFrame.put_MarginTop(0.);


   //Создаем таблицу бокового штампа на 2-й странице
   oRan = oTextFrame.get_TextRange();
   oTables = this->Word_blank.GetTables();

   //добавить таблицу в коллекцию
   oTable = oTables.Add(oRan,3,2,COleVariant(short(wdWord9TableBehavior)),COleVariant(short(wdAutoFitFixed)));
   //Установка направления текста в таблице
   oTable.Select();
   oSel.SetOrientation(wdTextOrientationUpward);

   //Установка шрифта в таблице
   oFont.SetSize(10);
   oFont.SetName("Arial");

   //Устанавливаем минимальные поля ячеек таблицы
	oTable.SetTopPadding(0.);
   oTable.SetBottomPadding(0.);
   oTable.SetLeftPadding(0.);
   oTable.SetRightPadding(0.);

   //Устанавливаем минимальное расстояние между ячейками таблицы
   oTable.SetSpacing(0.);
   //Устанавливаем автоподгонку размеров ячеек под содержимое
   oTable.SetAllowAutoFit(BOOL(true));

   //Устанавливаем вертикальное выравнивание в ячейках таблицы
   oTable.Select();
   oCells=oSel.GetCells();
   oCells.SetVerticalAlignment(wdCellAlignVerticalCenter);

   //Устанавливаем высоту строк
   oRows=oTable.GetRows();
   oRow=oRows.Item(1);
   oRow.SetHeight(25.*2.783, wdRowHeightExactly);
   oRow=oRows.Item(2);
   oRow.SetHeight(35.*2.783, wdRowHeightExactly);
   oRow=oRows.Item(3);
   oRow.SetHeight(25.*2.783, wdRowHeightExactly);

   //Устанавливаем ширину столбцов
   oColumns=oTable.GetColumns();
   oColumn=oColumns.Item(1);
   oColumn.SetPreferredWidth(5.*MM2PH);
   oColumn=oColumns.Item(2);
   oColumn.SetPreferredWidth(7.*MM2PH);

   //Устанавливаем толщину границ ячеек таблицы
   oBorders=oTable.GetBorders();
   oBorder=oBorders.Item(wdBorderLeft);
   oBorder.SetLineWidth(wdLineWidth150pt);
   oBorder=oBorders.Item(wdBorderTop);
   oBorder.SetLineWidth(wdLineWidth150pt);
   oBorder=oBorders.Item(wdBorderBottom);
   oBorder.SetLineWidth(wdLineWidth150pt);
   oBorder=oBorders.Item(wdBorderRight);
   oBorder.SetLineWidth(wdLineWidth150pt);
   oBorder=oBorders.Item(wdBorderHorizontal);
   oBorder.SetLineWidth(wdLineWidth150pt);
   oBorder=oBorders.Item(wdBorderVertical);
   oBorder.SetLineWidth(wdLineWidth150pt);

   //Проставляем надписи в графах
   //Графа "Инв. № подл."
   oCell = oTable.Cell(3,1);
   oRan = oCell.GetRange();
   oRan.SetText("Инв. № подл.");
   oCell.Select();
   oFont.SetSize(8);
   oSel.BoldRun();
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);

   //Графа 20
   oCell = oTable.Cell(3,2);
   oRan = oCell.GetRange();
   oRan.SetText(cDiD->sN_podl);
   oCell.Select();
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);

   //Графа "Подп. и дата"
   oCell = oTable.Cell(2,1);
   oRan = oCell.GetRange();
   oRan.SetText("Подп. и дата");
   oCell.Select();
   oFont.SetSize(8);
   oSel.BoldRun();
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);

   //Графа "Взам. инв. №"
   oCell = oTable.Cell(1,1);
   oRan = oCell.GetRange();
   oRan.SetText("Взам. инв. №");
   oCell.Select();
   oFont.SetSize(8);
   oSel.BoldRun();
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);

   //Графа 22
   oCell = oTable.Cell(1,2);
   oRan = oCell.GetRange();
   oRan.SetText(cDiD->sN_star_podl);
   oCell.Select();
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);


   //Создаем таблицу основного (маленького по Форме 6) штампа на 2-й странице
   View.put_SeekView(wdSeekCurrentPageHeader);
   oSel.MoveDown(COleVariant(short(wdLine)),COleVariant(long(2)),COleVariant(short(wdMove)));
   oRan = oSel.GetRange();
   oTables = this->Word_blank.GetTables();

   //добавить таблицу в коллекцию
   oTable = oTables.Add(oRan,4,8,COleVariant(short(wdWord9TableBehavior)),COleVariant(short(wdAutoFitFixed)));
   //Установка шрифта в таблице
   oTable.Select();
   oFont.SetSize(10);
   oFont.SetName("Arial");

   //Устанавливаем минимальные поля ячеек таблицы
	oTable.SetTopPadding(0.);
   oTable.SetBottomPadding(0.);
   oTable.SetLeftPadding(0.);
   oTable.SetRightPadding(0.);

   //Устанавливаем минимальное расстояние между ячейками таблицы
   oTable.SetSpacing(0.);
   //Устанавливаем автоподгонку размеров ячеек под содержимое
   oTable.SetAllowAutoFit(BOOL(true));

   //Устанавливаем вертикальное выравнивание в ячейках таблицы
   oTable.Select();
   oCells=oSel.GetCells();
   oCells.SetVerticalAlignment(wdCellAlignVerticalCenter);

   //Устанавливаем высоту строк
   oRows=oTable.GetRows();
   oRows.SetHeight(5.*2.771, wdRowHeightExactly);
   oRow=oRows.Item(2);
   oRow.SetHeight(2.*2.771, wdRowHeightExactly);
   oRow=oRows.Item(3);
   oRow.SetHeight(3.*2.771, wdRowHeightExactly);

   //Устанавливаем ширину столбцов
   oColumns=oTable.GetColumns();
   oColumn=oColumns.Item(1);
   oColumn.SetPreferredWidth(10.*MM2PH);
   oColumn=oColumns.Item(2);
   oColumn.SetPreferredWidth(10.*MM2PH);
   oColumn=oColumns.Item(3);
   oColumn.SetPreferredWidth(10.*MM2PH);
   oColumn=oColumns.Item(4);
   oColumn.SetPreferredWidth(10.*MM2PH);
   oColumn=oColumns.Item(5);
   oColumn.SetPreferredWidth(15.*MM2PH);
   oColumn=oColumns.Item(6);
   oColumn.SetPreferredWidth(10.*MM2PH);
   oColumn=oColumns.Item(7);
   oColumn.SetPreferredWidth(110.*MM2PH);
   oColumn=oColumns.Item(8);
   oColumn.SetPreferredWidth(10.*MM2PH);

   //Объединяем ячейки
   oCell = oTable.Cell(2,1);
   oCell.Select();
   oSel.MoveDown(COleVariant(short(wdLine)),COleVariant(long(1)),COleVariant(short(wdExtend)));
   oCells=oSel.GetCells();
   oCells.Merge();
   oCell = oTable.Cell(2,2);
   oCell.Select();
   oSel.MoveDown(COleVariant(short(wdLine)),COleVariant(long(1)),COleVariant(short(wdExtend)));
   oCells=oSel.GetCells();
   oCells.Merge();
   oCell = oTable.Cell(2,3);
   oCell.Select();
   oSel.MoveDown(COleVariant(short(wdLine)),COleVariant(long(1)),COleVariant(short(wdExtend)));
   oCells=oSel.GetCells();
   oCells.Merge();
   oCell = oTable.Cell(2,4);
   oCell.Select();
   oSel.MoveDown(COleVariant(short(wdLine)),COleVariant(long(1)),COleVariant(short(wdExtend)));
   oCells=oSel.GetCells();
   oCells.Merge();
   oCell = oTable.Cell(2,5);
   oCell.Select();
   oSel.MoveDown(COleVariant(short(wdLine)),COleVariant(long(1)),COleVariant(short(wdExtend)));
   oCells=oSel.GetCells();
   oCells.Merge();
   oCell = oTable.Cell(2,6);
   oCell.Select();
   oSel.MoveDown(COleVariant(short(wdLine)),COleVariant(long(1)),COleVariant(short(wdExtend)));
   oCells=oSel.GetCells();
   oCells.Merge();
   oCell = oTable.Cell(1,7);
   oCell.Select();
   oSel.MoveDown(COleVariant(short(wdLine)),COleVariant(long(3)),COleVariant(short(wdExtend)));
   oCells=oSel.GetCells();
   oCells.Merge();
   oCell = oTable.Cell(1,8);
   oCell.Select();
   oSel.MoveDown(COleVariant(short(wdLine)),COleVariant(long(1)),COleVariant(short(wdExtend)));
   oCells=oSel.GetCells();
   oCells.Merge();
   oCell = oTable.Cell(3,8);
   oCell.Select();
   oSel.MoveDown(COleVariant(short(wdLine)),COleVariant(long(1)),COleVariant(short(wdExtend)));
   oCells=oSel.GetCells();
   oCells.Merge();

   //Устанавливаем толщину границ ячеек таблицы
   oBorders=oTable.GetBorders();
   oBorder=oBorders.Item(wdBorderLeft);
   oBorder.SetLineWidth(wdLineWidth150pt);
   oBorder=oBorders.Item(wdBorderTop);
   oBorder.SetLineWidth(wdLineWidth150pt);
   oBorder=oBorders.Item(wdBorderBottom);
   oBorder.SetLineWidth(wdLineWidth150pt);
   oBorder=oBorders.Item(wdBorderRight);
   oBorder.SetLineWidth(wdLineWidth150pt);
   oBorder=oBorders.Item(wdBorderHorizontal);
   oBorder.SetLineWidth(wdLineWidth150pt);
   oBorder=oBorders.Item(wdBorderVertical);
   oBorder.SetLineWidth(wdLineWidth150pt);
   oCell = oTable.Cell(1,1);
   oCell.Select();
   oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(5)),COleVariant(short(wdExtend)));
   oBorders=oSel.GetBorders();
   oBorder=oBorders.Item(wdBorderBottom);
   oBorder.SetLineWidth(wdLineWidth075pt);

   //Проставляем надписи в графах
   //Графа "Изм."
   oCell = oTable.Cell(4,1);
   oRan = oCell.GetRange();
   oRan.SetText("Изм.");
   oCell.Select();
   oFont.SetSize(7);
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);

   //Графа "Кол. уч."
   oCell = oTable.Cell(4,2);
   oRan = oCell.GetRange();
   oRan.SetText("Кол. уч.");
   oCell.Select();
   oFont.SetSize(7);
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);

   //Графа "Лист" в таблице изменений
   oCell = oTable.Cell(4,3);
   oRan = oCell.GetRange();
   oRan.SetText("Лист");
   oCell.Select();
   oFont.SetSize(7);
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);

   //Графа "№ док."
   oCell = oTable.Cell(4,4);
   oRan = oCell.GetRange();
   oRan.SetText("№ док.");
   oCell.Select();
   oFont.SetSize(7);
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);

   //Графа "Подп." в таблице изменений
   oCell = oTable.Cell(4,5);
   oRan = oCell.GetRange();
   oRan.SetText("Подп.");
   oCell.Select();
   oFont.SetSize(7);
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);

   //Графа "Дата" в таблице изменений
   oCell = oTable.Cell(4,6);
   oRan = oCell.GetRange();
   oRan.SetText("Дата");
   oCell.Select();
   oFont.SetSize(7);
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);

   //Графа 1
   oCell = oTable.Cell(1,7);
   oRan = oCell.GetRange();
   oRan.SetText(cDiD->sObozn_doc);
   oCell.Select();
   oFont.SetSize(14);
   oSel.BoldRun();
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);

   //Графа "Лист"
   oCell = oTable.Cell(1,8);
   oRan = oCell.GetRange();
   oRan.SetText("Лист");
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);

   //Графа 7
   oCell = oTable.Cell(3,8);
   oCell.Select();
   oRan = oSel.GetRange();
   oSel.Collapse(COleVariant(short(wdCollapseStart)));
   oSel.SetText(sN_list_pref);
   oSel.Collapse(COleVariant(short(wdCollapseEnd)));
   oFields=oSel.GetFields();
	oFields.Add(oSel.GetRange(), COleVariant(short(wdFieldPage)), covOptional, covOptional);//Вставка поля "номер страницы"
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);


   //Рисуем прямоугольную рамку на второй странице
   oShape=oShapes.AddShape(msoShapeRectangle,20.*2.82,5.*2.787,185.*2.834,287.*2.787,covOptional);
   oFillFormat=oShape.get_Fill();
   oLineFormat=oShape.get_Line();
   oFillFormat.put_Visible(long (FALSE));
   oLineFormat.put_Weight(0.5*MM2PH);


   //Удаляем 2-ю страницу
   View.put_SeekView(wdSeekMainDocument);//Выйти из колонтитула
   oSel.MoveUp(COleVariant(short(wdParagraph)),COleVariant(short(1)),COleVariant(short(wdExtend)));
   oSel.Delete(COleVariant(short(wdCharacter)),COleVariant(short(1)));

   return TRUE;
}//BOOL	cBlank_A4_f5::Draw_blank(cData_interval_DMR &cDiD)

