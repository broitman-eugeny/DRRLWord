//Функции класса отчета по результатам расчета качественных показателей otchet_DMR
#include <afx.h>
#include "afxdisp.h"//Необходима для работы AfxOleInit()
#include "msword9.h"//Заголовочный файл, полученный с помощью ClassWizard Visual Studio
#include "DMRWord.h"//Заголовочный файл с описаниями классов MS Word



//Пустой конструктор. Используется когда данные для отчета берутся из объекта типа cData_interval_DMR
otchet_DMR::otchet_DMR()
{
}

//Отображает отчет результатов расчета качественных показателей ЦРРЛ в краткой или полной форме.
//Параметры берет по ссылке на объект типа cData_interval_DMR, описанный в "DMR.h"
//В случае успеха возвращает TRUE, иначе - FALSE.
BOOL otchet_DMR::Draw_otchet_DMR(cData_interval_DMR *cDiD)
{
   COleVariant  covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);

   if(!Draw_blank(cDiD)) //Если не удалось отобразить бланк с рамкой и основной надписью
   {
      //Выгрузка Word
      app.Quit(COleVariant(short(wdDoNotSaveChanges)), COleVariant(short(NULL/*wdPromptUser*/)), COleVariant(short(false)));
   	return FALSE;
   }

   //Установка параметров страницы
   CPageSetup oPageSetup;
   oPageSetup=Word_blank.GetPageSetup();
   oPageSetup.put_TopMargin(5.0*MM2PV);//Расстояние от верхнего края листа
   Word_blank.SetPageSetup(oPageSetup);

   //Установка шрифта вне таблицы
   _Font oFont;
   Selection oSel;
   oSel = app.GetSelection();
   oFont=oSel.GetFont();
   oFont.SetSize(1);
   oFont.SetName("Arial");


   //Создаем шапку таблицы отчета
   Tables oTables;
   Table oTable;
   Range oRan;
   oRan = oSel.GetRange();
   oTables = this->Word_blank.GetTables();
   //добавить таблицу (1 строка, 5 столбцов) в коллекцию
   oTable = oTables.Add(oRan,1,5,COleVariant(short(wdWord9TableBehavior)),COleVariant(short(wdAutoFitFixed)));
   //Положение таблицы на странице
   Rows oRows;
   oRows=oTable.GetRows();
   oRows.SetLeftIndent(0.1);

   //Установка шрифта в таблице
   oTable.Select();
   oFont.SetSize(8);
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
   Row oRow;
   oRow=oRows.Item(1);
   oRow.SetHeight(4.2*MM2PV);

   //Устанавливаем ширину столбцов
   Columns oColumns;
   oColumns=oTable.GetColumns();
   Column oColumn;
   oColumn=oColumns.Item(1);
   oColumn.SetWidth(10.*MM2PH, wdAdjustSameWidth);
   oColumn=oColumns.Item(2);
   oColumn.SetWidth(108.9*MM2PH,wdAdjustSameWidth);
   oColumn=oColumns.Item(3);
   oColumn.SetWidth(25.*MM2PH, wdAdjustSameWidth);
   oColumn=oColumns.Item(4);
   oColumn.SetWidth(20.*MM2PH, wdAdjustSameWidth);
   oColumn=oColumns.Item(5);
   oColumn.SetWidth(20.*MM2PH, wdAdjustSameWidth);

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
   oBorder=oBorders.Item(wdBorderVertical);
   oBorder.SetLineWidth(wdLineWidth150pt);

   //Проставляем надписи в графах
   //Графа "№ п/п"
   Cell oCell;
   oCell = oTable.Cell(1,1);
   oRan = oCell.GetRange();
   oRan.SetText("№ п/п");
   //Графа "Наименование параметра"
   oCell = oTable.Cell(1,2);
   oRan = oCell.GetRange();
   oRan.SetText("Наименование параметра");

   //Графа "Обозначение"
   oCell = oTable.Cell(1,3);
   oRan = oCell.GetRange();
   oRan.SetText("Обозначение");
   //Графа "Размерн."
   oCell = oTable.Cell(1,4);
   oRan = oCell.GetRange();
   oRan.SetText("Размерн.");
   //Графа "Значение"
   oCell = oTable.Cell(1,5);
   oRan = oCell.GetRange();
   oRan.SetText("Значение");
   //Выравнивание текста по центру
   oTable.Select();
   Paragraphs oPars;
   oPars=oSel.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);


   //Добавляем строку
   oSel.SelectRow();
   oSel.InsertRowsBelow(COleVariant(short(1)));


   //Вставляем в первый столбец нумерованный список
   CListGalleries oLGs;
   oLGs=app.GetListGalleries();
   CListGallery oLG;
   oLG=oLGs.Item(wdNumberGallery);
   CListTemplates oLTs;
   oLTs=oLG.get_ListTemplates();
   CListTemplate oLT;
   oLT=oLTs.Item(COleVariant(short(5)));
   CListLevels oLLs;
   oLLs=oLT.get_ListLevels();
   CListLevel oLL;
   oLL=oLLs.Item(1);
   oLL.put_NumberFormat("%1");
   oLL.put_NumberPosition(5.0*MM2PH);
   oLL.put_ResetOnHigher((long)0);
   oLL.put_StartAt((long)1);
   oCell = oTable.Cell(2,1);
   oCell.Select();
   oRan = oCell.GetRange();
   CListFormat oLF;
   oLF=oRan.GetListFormat();
   oLF.ApplyListTemplate(oLT, COleVariant(short(False)), COleVariant(short(wdListApplyToWholeList)), COleVariant(short(wdWord10ListBehavior)));


   //Устанавливаем выравнивание по левому краю в графе "Наименование параметра"
   oCell = oTable.Cell(2,2);
   oCell.Select();
   oPars=oSel.GetParagraphs();
   oPars.SetAlignment(AL_LEFT);

   //Добавляем строку
   oSel.SelectRow();
   oSel.InsertRowsBelow(COleVariant(short(1)));


   //Установка шрифта в столбце "Обозначение"
   oCell = oTable.Cell(3,3);
   oCell.Select();
   oFont.SetSize(10);
   oFont.SetName("Times New Roman");


   //Устанавливаем шапку таблицы в "Заголовок"
   oRow=oRows.Item(1);
   oRow.SetHeadingFormat(True);

   int N_strok=0;//Счетчик заполненных строк не включая шапку

   CParagraphFormat oParFormat;//для форматирования выравнивания и вставки табуляторов


   //Полный отчет
   if(cDiD->Vid_otcheta==POLNY)
   {


      //0	Дата и время расчета
      //Объединяем ячейки
      oCell = oTable.Cell(N_strok+2,3);
   	oCell.Select();
   	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdExtend)));
   	oCells=oSel.GetCells();
   	oCells.Merge();
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
   	oRan = oCell.GetRange();
   	oRan.SetText(cDiD->sCurDate[0]);

      //Заполняем графы "Обозначение", "Размерн.", "Значение"
      oCell = oTable.Cell(N_strok+2,3);
   	oRan = oCell.GetRange();
   	oRan.SetText(cDiD->sCurDate[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //1	Путь и имя файла профиля
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sFileName[0]);
      //Заполняем графы "Обозначение", "Размерн.", "Значение"
      oCell = oTable.Cell(N_strok+2,3);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sFileName[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;

      //2	Дата и время последнего изменения файла профиля
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sFileDate[0]);
      //Заполняем графы "Обозначение", "Размерн.", "Значение"
      oCell = oTable.Cell(N_strok+2,3);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sFileDate[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;

      //3	Название станции слева из файла профиля
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sLeftStationName[0]);
      //Заполняем графы "Обозначение", "Размерн.", "Значение"
      oCell = oTable.Cell(N_strok+2,3);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sLeftStationName[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;

      //4	Название станции справа из файла профиля
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sRightStationName[0]);
      //Заполняем графы "Обозначение", "Размерн.", "Значение"
      oCell = oTable.Cell(N_strok+2,3);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sRightStationName[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;

      //5	град. мин., Прямой азимут из файла профиля
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sForwardAzimuth[0]);
      //Заполняем графы "Обозначение", "Размерн.", "Значение"
      oCell = oTable.Cell(N_strok+2,3);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sForwardAzimuth[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;

      //6	град. мин., Обратный азимут из файла профиля
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sBackwardAzimuth[0]);
      //Заполняем графы "Обозначение", "Размерн.", "Значение"
      oCell = oTable.Cell(N_strok+2,3);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sBackwardAzimuth[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;

      //7	м, Отметка рельефа станции слева из файла профиля
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sr_0[0]);
      //Заполняем графы "Обозначение", "Размерн.", "Значение"
      oCell = oTable.Cell(N_strok+2,3);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sr_0[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;

      //8	м, Отметка рельефа станции справа из файла профиля
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sr_R[0]);
      //Заполняем графы "Обозначение", "Размерн.", "Значение"
      oCell = oTable.Cell(N_strok+2,3);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sr_R[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //9	Тип оборудования
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sTyp_obor[0]);

      //Заполняем графы "Обозначение", "Размерн.", "Значение"
      oCell = oTable.Cell(N_strok+2,3);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sTyp_obor[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //10	Назначение РРЛ (Протяженность линии, км)
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sNaznachenie[0]);
      //Заполняем графы "Обозначение", "Размерн.", "Значение"
      oCell = oTable.Cell(N_strok+2,3);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sNaznachenie[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //11	Линия реконструируемая или вновь проектируемая
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sReconstr[0]);
      //Заполняем графы "Обозначение", "Размерн.", "Значение"
      oCell = oTable.Cell(N_strok+2,3);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sReconstr[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //12	Характер трассы(по влажности, по пересеченности, по высоте рельефа)
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sHarakter_trassy[0]);
      //Заполняем графы "Обозначение", "Размерн.", "Значение"
      oCell = oTable.Cell(N_strok+2,3);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sHarakter_trassy[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      if(cDiD->sPol[3]!="")
      {
      	//13	Поляризация
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oRan = oCell.GetRange();
      	oRan.SetText(cDiD->sPol[0]);
      	//Заполняем графы "Обозначение", "Размерн.", "Значение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oRan = oCell.GetRange();
      	oRan.SetText(cDiD->sPol[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      //14	Тип системы модуляции
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sTyp_sys[0]);
      //Заполняем графы "Обозначение", "Размерн.", "Значение"
      oCell = oTable.Cell(N_strok+2,3);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sTyp_sys[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //15	км, Протяженность интервала
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sR[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sR[1]);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sR[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sR[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //16	ГГц, Частотный диапазон
      //Добавляем строку
      oCell.Select();
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sf[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sf[1]);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sf[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sf[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //17	м, Длина волны
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sLambda[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
   	oSel.InsertSymbol((256*(unsigned char)(cDiD->sLambda[1][0])+(unsigned char)(cDiD->sLambda[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sLambda[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sLambda[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //18	Мбит/с, Скорость передачи цифрового потока
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sC[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sC[1]);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sC[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sC[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //19	дБм, Мощность передатчика
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sP_pd[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sP_pd[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
   	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sP_pd[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sP_pd[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //21	дБм, Паспортная пороговая чувствительность приемника при заданном K_osh
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sP_pm_por_K_osh[0]);
      //Добавляем табулятор
      	oParFormat=oSel.GetParagraphFormat();
      	CTabStops oTabStops;
      	oTabStops=oParFormat.get_TabStops();
      	oTabStops.Add(80.2*MM2PH, COleVariant(short(wdAlignTabLeft)), COleVariant(short(wdTabLeaderSpaces)));
      	oParFormat.put_TabStops(oTabStops);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.SetText(cDiD->sK_osh[3]);//20	Коэффициент ошибок для которого приводится паспортная пороговая чувствительность приемника

      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sP_pm_por_K_osh[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
   	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sP_pm_por_K_osh[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sP_pm_por_K_osh[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //23	дБм, Пороговая чувствительность приемника при BER
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sP_pm_por[0]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.SetText(cDiD->sBER[3]);//22	Коэффициент ошибок по битам, в зависимости от скорости передачи цифрового потока
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sP_pm_por[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
   	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sP_pm_por[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sP_pm_por[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //24	Наличие пассивной ретрансляции сигнала на трассе (есть/нет)
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sPass_retr[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sPass_retr[1]);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sPass_retr[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sPass_retr[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //25	Тип АМС слева (Трубчатая опора/Решетчатая или железобетонная опора)
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sTruba_pd[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sTruba_pd[1]);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sTruba_pd[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sTruba_pd[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //26	Тип АМС справа (Трубчатая опора/Решетчатая или железобетонная опора)
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sTruba_pm[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sTruba_pm[1]);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sTruba_pm[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sTruba_pm[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //27	Антенная система слева - перископическая? (да/нет)
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sPeriskop_pd[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sPeriskop_pd[1]);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sPeriskop_pd[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sPeriskop_pd[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //28	Антенная система справа - перископическая? (да/нет)
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sPeriskop_pm[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sPeriskop_pm[1]);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sPeriskop_pm[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sPeriskop_pm[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //30	дБи, Коэффициент усиления основной антенны слева
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.SetText(cDiD->sG_pd[0]);
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.InsertSymbol(216, COleVariant("Arial"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      oSel.SetText(cDiD->sD_pd[3]);//29	м, Диаметр основной антенны слева
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sG_pd[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
   	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sG_pd[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sG_pd[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //32	дБи, Коэффициент усиления основной антенны справа
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.SetText(cDiD->sG_pm[0]);
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.InsertSymbol(216, COleVariant("Arial"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      oSel.SetText(cDiD->sD_pm[3]);//31	м, Диаметр основной антенны справа
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sG_pm[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
   	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sG_pm[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sG_pm[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      if(cDiD->sG1_dop[3]!="")
      {
      	//215	дБи, Коэффициент усиления дополнительной антенны слева (номинальное значение/с учетом ограничений)
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.SetText(cDiD->sG1_dop[0]);
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.InsertSymbol(216, COleVariant("Arial"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText(cDiD->sD1_dop[3]);//214	м, Диаметр дополнительной антенны слева
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sG1_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sG1_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sG1_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sG2_dop[3]!="")
      {
      	//217	дБи, Коэффициент усиления дополнительной антенны справа (номинальное значение/с учетом ограничений)
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.SetText(cDiD->sG2_dop[0]);
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.InsertSymbol(216, COleVariant("Arial"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText(cDiD->sD2_dop[3]);//216	м, Диаметр дополнительной антенны справа
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sG2_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sG2_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sG2_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      //33	дБ, Постоянные потери в АФТ слева
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sEta_post_pd[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
   	oSel.InsertSymbol((256*(unsigned char)(cDiD->sEta_post_pd[1][0])+(unsigned char)(cDiD->sEta_post_pd[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      oSel.SetText((cDiD->sEta_post_pd[1]).Right(10));
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
   	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sEta_post_pd[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sEta_post_pd[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //34	дБ, Постоянные потери в АФТ справа
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sEta_post_pm[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
   	oSel.InsertSymbol((256*(unsigned char)(cDiD->sEta_post_pm[1][0])+(unsigned char)(cDiD->sEta_post_pm[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      oSel.SetText((cDiD->sEta_post_pm[1]).Right(11));
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
   	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sEta_post_pm[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sEta_post_pm[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      if(cDiD->sAlfa_AVT_pd[3]!="")
      {
      	//35	дБ/м, Погонное затухание волновода основной антенны слева
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sAlfa_AVT_pd[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.InsertSymbol((256*(unsigned char)(cDiD->sAlfa_AVT_pd[1][0])+(unsigned char)(cDiD->sAlfa_AVT_pd[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sAlfa_AVT_pd[1]).Right(9));
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sAlfa_AVT_pd[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sAlfa_AVT_pd[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sAlfa_AVT_pm[3]!="")
      {
      	//36	дБ/м, Погонное затухание волновода основной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sAlfa_AVT_pm[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.InsertSymbol((256*(unsigned char)(cDiD->sAlfa_AVT_pm[1][0])+(unsigned char)(cDiD->sAlfa_AVT_pm[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sAlfa_AVT_pm[1]).Right(10));
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sAlfa_AVT_pm[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sAlfa_AVT_pm[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sL_AVT_pd[3]!="")
      {
      	//37	м, Длина волновода основной антенны слева
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sL_AVT_pd[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sL_AVT_pd[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sL_AVT_pd[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sL_AVT_pd[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sL_AVT_pm[3]!="")
      {
      	//38	м, Длина волновода основной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sL_AVT_pm[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sL_AVT_pm[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sL_AVT_pm[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sL_AVT_pm[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      //39	м, Высота центра раскрыва основной антенны слева
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sH1[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sH1[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
   	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sH1[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sH1[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //40	м, Высота центра раскрыва основной антенны справа
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sH2[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sH2[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
   	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sH2[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sH2[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      if(cDiD->sH1_dop[3]!="")
      {
      	//212	м, Высота центра раскрыва дополнительной антенны слева
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sH2_dop[3]!="")
      {
      	//213	м, Высота центра раскрыва дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      //41	м, Средняя высота трассы луча над уровнем моря
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sh_sr[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sh_sr[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
   	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sh_sr[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sh_sr[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //42	Номер климатического района
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sRaion[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sRaion[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
   	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sRaion[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sRaion[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //43	1/м, Среднее значение градиента диэл. проницаемости воздуха
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sg_sr[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sg_sr[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
   	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sg_sr[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sg_sr[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //44	1/м, Стандартное отклонение градиента диэл. проницаемости воздуха
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sSigma[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.InsertSymbol((256*(unsigned char)(cDiD->sSigma[1][0])+(unsigned char)(cDiD->sSigma[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sSigma[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sSigma[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      if(cDiD->sSigma_ot_R[3]!="")
      {
      	//45	1/м, СКО градиента диэл. проницаемости воздуха в зависимости от расстояния
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sSigma_ot_R[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sSigma_ot_R[1][0])+(unsigned char)(cDiD->sSigma_ot_R[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sSigma_ot_R[1]).Right(3));
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sSigma_ot_R[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sSigma_ot_R[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      //46	дБ, Суммарные потери на интервале в антенно-фидерных трактах
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sEta_AFT[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.InsertSymbol((256*(unsigned char)(cDiD->sEta_AFT[1][0])+(unsigned char)(cDiD->sEta_AFT[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      oSel.SetText((cDiD->sEta_AFT[1]).Right(3));
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
   	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sEta_AFT[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sEta_AFT[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //47	дБ, Ослабление сигнала в свободном пространстве
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sW_0[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sW_0[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
   	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sW_0[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sW_0[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      if(cDiD->sV_difr_sr[3]!="")
      {
      	//48	дБ, Среднее ослабление за счет дифракции (при средней рефракции)
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_difr_sr[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_difr_sr[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_difr_sr[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_difr_sr[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sVlazh_para[3]!="")
      {
      	//49	г/м3, Влажность пара
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sVlazh_para[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sVlazh_para[1]);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sVlazh_para[2]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(3)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSuperscript(True);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sVlazh_para[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sV_v_gazah[3]!="")
      {
      	//50	дБ, Среднее ослабление сигнала в газах атмосферы
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_v_gazah[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_v_gazah[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_v_gazah[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_v_gazah[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sV_50_proc[3]!="")
      {
      	//51	дБ, Среднее ослабление сигнала из-за перепада высот на горных и высокогорных трассах
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_50_proc[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_50_proc[1]);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_50_proc[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_50_proc[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      //52	дБм, Средний уровень сигнала на входе приемника
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sP_pm[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sP_pm[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
   	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sP_pm[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sP_pm[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      if(cDiD->sDelta_V_tip[3]!="")
      {
      	//53	дБ, Поправка на Vmin для типовых параметров аппаратуры
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_tip[0]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(13)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(3)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_V_tip[1][0])+(unsigned char)(cDiD->sDelta_V_tip[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sDelta_V_tip[1]).Right(5));
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_tip[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_tip[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sV_0_min_d[3]!="")
      {
      	//54	дБ, Минимально допустимый множитель ослабления по дождям без учета деградации
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_0_min_d[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_0_min_d[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_0_min_d[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_0_min_d[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sV_0_min_0[3]!="")
      {
      	//55	дБ, Минимально допустимый множитель ослабления по субрефракции без учета деградации
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_0_min_0[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_0_min_0[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_0_min_0[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_0_min_0[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sZ_por_dop[3]!="")
      {
      	//56	дБ, Пороговое отношение (помеха/Рпм пор, при деградации порогового уровня на 3 дБ)
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_por_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_por_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_por_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_por_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sZ_por_dop_s[3]!="")
      {
      	//57	дБ, Отношение мощности мешающего сигнала соседнего ствола к мощности полезного сигнала, вызывающее в канале Pош_макс при деградации порогового уровня на 3 дБ
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_por_dop_s[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_por_dop_s[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_por_dop_s[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_por_dop_s[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sZ_sosedn_kanal[3]!="")
      {
      	//58	дБ, Превышение отношением (помеха/Рпм пор, при деградации порогового уровня на 3 дБ) порогового отношения (помехи от соседнего канала/Рпм пор)
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_sosedn_kanal[0]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(12)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(8)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(23)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(10)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_sosedn_kanal[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_sosedn_kanal[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_sosedn_kanal[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sDelta_V_sosedn_kanal[3]!="")
      {
      	//59	дБ, Деградация порогового уровня из-за влияния помех от соседнего канала
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_sosedn_kanal[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_V_sosedn_kanal[1][0])+(unsigned char)(cDiD->sDelta_V_sosedn_kanal[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sDelta_V_sosedn_kanal[1]).Right(2));
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_sosedn_kanal[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_sosedn_kanal[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      //60	Наличие "Co-channel" (есть/нет)
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sCochannel[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sCochannel[1]);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sCochannel[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sCochannel[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      if(cDiD->sD_p_0[3]!="")
      {
      	//61	дБ, Коэффициент поляризационной защиты при отсутствии замираний основной приемной антенны
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sD_p_0[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sD_p_0[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sD_p_0[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sD_p_0[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sQ_p_cochannel[3]!="")
      {
      	//62	дБ, Коэффициент, учитывающий наклон кроссполяризационной диаграммы направленности антенны
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sQ_p_cochannel[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sQ_p_cochannel[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sQ_p_cochannel[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sQ_p_cochannel[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sD_p[3]!="")
      {
      	//63	дБ, Коэффициент поляризационной защиты в условиях интерференционных замираний
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sD_p[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sD_p[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sD_p[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sD_p[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sD_p_dojd[3]!="")
      {
      	//64	дБ, Коэффициент поляризационной защиты, обусловленный влиянием дождей
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sD_p_dojd[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sD_p_dojd[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sD_p_dojd[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sD_p_dojd[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sI_xpic[3]!="")
      {
      	//65	дБ, Выигрыш компенсатора кроссполяризационных помех
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sI_xpic[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sI_xpic[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sI_xpic[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sI_xpic[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sZ_polyariz_cochannel[3]!="")
      {
      	//66	дБ, Превышение отношением (кроссполяризационной помехи/Рпм пор) порогового отношения (помеха/Рпм пор, при деградации порогового уровня на 3 дБ), для учета в расчетах Тинт (в расчетах Т0 кроссполяризационная помеха не учитывается)
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_polyariz_cochannel[0]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(12)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(16)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(7)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(18)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(4)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_polyariz_cochannel[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_polyariz_cochannel[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_polyariz_cochannel[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sZ_polyariz_cochannel_dojd[3]!="")
      {
      	//67	дБ, Превышение отношением (кроссполяризационной помехи/Рпм пор) порогового отношения (помеха/Рпм пор, при деградации порогового уровня на 3 дБ), для учета в расчетах Тд (в расчетах Т0 кроссполяризационная помеха не учитывается)
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_polyariz_cochannel_dojd[0]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(12)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(16)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(7)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(18)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(4)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_polyariz_cochannel_dojd[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_polyariz_cochannel_dojd[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_polyariz_cochannel_dojd[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sDelta_V_polyariz_cochannel[3]!="")
      {
      	//68	дБ, Деградация порогового уровня из-за влияния кроссполяризационных помех для учета в расчетах Тинт (в расчетах Т0 кроссполяризационная помеха не учитывается)
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_polyariz_cochannel[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_V_polyariz_cochannel[1][0])+(unsigned char)(cDiD->sDelta_V_polyariz_cochannel[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sDelta_V_polyariz_cochannel[1]).Right(7));
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_polyariz_cochannel[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_polyariz_cochannel[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sDelta_V_polyariz_cochannel_dojd[3]!="")
      {
      	//69	дБ, Деградация порогового уровня из-за влияния кроссполяризационных помех для учета в расчетах Тд (в расчетах Т0 кроссполяризационная помеха не учитывается)
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_polyariz_cochannel_dojd[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_V_polyariz_cochannel_dojd[1][0])+(unsigned char)(cDiD->sDelta_V_polyariz_cochannel_dojd[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sDelta_V_polyariz_cochannel_dojd[1]).Right(5));
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_polyariz_cochannel_dojd[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_polyariz_cochannel_dojd[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sN_mesh[3]!="")
      {
      	//70	Количество мешающих интервалов, работающих на частоте полезного сигнала
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sN_mesh[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sN_mesh[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sN_mesh[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sN_mesh[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sR_mesh[3]!="")
      {
      	//71	км, Протяженности мешающих интервалов (через слэш)
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_mesh[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_mesh[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_mesh[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_mesh[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sF_ot_alfa[3]!="")
      {
      	//72	дБ, Ослабления мешающих сигналов за счет диаграмм направленности антенн (через слэш)
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sF_ot_alfa[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.SetText((cDiD->sF_ot_alfa[1]).Left(2));
      	oSel.Collapse(COleVariant(short(wdCollapseEnd)));
         oSel.InsertSymbol((256*(unsigned char)(cDiD->sF_ot_alfa[1][2])+(unsigned char)(cDiD->sF_ot_alfa[1][3])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
         oSel.SetText((cDiD->sF_ot_alfa[1]).Right(1));
         //Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sF_ot_alfa[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sF_ot_alfa[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sP_pd_mesh[3]!="")
      {
      	//73	дБм, Мощности передатчиков мешающих интервалов (через слэш)
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sP_pd_mesh[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sP_pd_mesh[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sP_pd_mesh[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sP_pd_mesh[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sARM[3]!="")
      {
      	//74	дБ, Диапазон автоматической регулировки мощности передатчиков мешающих интервалов (через слэш)
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sARM[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sARM[1]);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sARM[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sARM[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sV_sr_mesh[3]!="")
      {
      	//75	дБ, Средние ослабления на трассах мешающих интервалов (через слэш)
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_sr_mesh[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sV_sr_mesh[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_sr_mesh[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_sr_mesh[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sG_pd_mesh[3]!="")
      {
      	//76	дБи, Коэффициенты усиления передающих антенн мешающих интервалов (через слэш)
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sG_pd_mesh[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sG_pd_mesh[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sG_pd_mesh[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sG_pd_mesh[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sL_AVT_pd_mesh[3]!="")
      {
      	//77	м, Длины волноводов передатчиков мешающих интервалов (через слэш)
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sL_AVT_pd_mesh[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sL_AVT_pd_mesh[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sL_AVT_pd_mesh[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sL_AVT_pd_mesh[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sZ_por[3]!="")
      {
      	//78	дБ, Отношение (суммарной помехи обратного направления и узлообразования/Рпм пор)
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_por[0]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(67)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sZ_por[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_por[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_por[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sZ_obr_uzl[3]!="")
      {
      	//79	дБ, Превышение отношением (суммарной помехи обратного направления и узлообразования/Рпм пор) порогового отношения (помеха/Рпм пор, при деградации порогового уровня на 3 дБ)
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_obr_uzl[0]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(12)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(3)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(16)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sZ_obr_uzl[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_obr_uzl[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_obr_uzl[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sDelta_V_obr_uzl[3]!="")
      {
      	//80	дБ, Деградация порогового уровня из-за влияния помех с обратного направления и узлообразования
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_obr_uzl[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_V_obr_uzl[1][0])+(unsigned char)(cDiD->sDelta_V_obr_uzl[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
         oSel.SetText((cDiD->sDelta_V_obr_uzl[1]).Right(9));
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_obr_uzl[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_obr_uzl[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sDelta_V_degr_int[3]!="")
      {
      	//81	дБ, Деградация порогового уровня из-за влияния помех на Tинт
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_degr_int[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_V_degr_int[1][0])+(unsigned char)(cDiD->sDelta_V_degr_int[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
         oSel.SetText((cDiD->sDelta_V_degr_int[1]).Right(10));
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_degr_int[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_degr_int[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sDelta_V_degr_subrefr[3]!="")
      {
      	//82	дБ, Деградация порогового уровня из-за влияния помех на T0
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_degr_subrefr[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_V_degr_subrefr[1][0])+(unsigned char)(cDiD->sDelta_V_degr_subrefr[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
         oSel.SetText((cDiD->sDelta_V_degr_subrefr[1]).Right(11));
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_degr_subrefr[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_degr_subrefr[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sDelta_V_degr_dojd[3]!="")
      {
      	//83	дБ, Деградация порогового уровня из-за влияния помех на Tд
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_degr_dojd[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_V_degr_dojd[1][0])+(unsigned char)(cDiD->sDelta_V_degr_dojd[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
         oSel.SetText((cDiD->sDelta_V_degr_dojd[1]).Right(11));
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_degr_dojd[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_degr_dojd[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sV_min_int[3]!="")
      {
      	//84	дБ, Минимально допустимый множитель интерференционного ослабления c учетом деградации порогового уровня, средних ослаблений и поправки на типовые параметры оборудования
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_int[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sV_min_int[1]);
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_int[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_int[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sV_min_subrefr[3]!="")
      {
      	//85	дБ, Минимально допустимый множитель субрефракционного ослабления c учетом деградации порогового уровня, средних ослаблений и поправки на типовые параметры оборудования
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_subrefr[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sV_min_subrefr[1]);
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_subrefr[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_subrefr[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sV_min_d[3]!="")
      {
      	//86	дБ, Минимально допустимый множитель ослабления в дожде c учетом деградации порогового уровня, средних ослаблений и поправки на типовые параметры оборудования
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_d[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sV_min_d[1]);
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_d[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_d[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sEkv[3]!="")
      {
      	//87	Наличие в системе эквалайзера (есть/нет)
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sEkv[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sEkv[1]);
         //Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sEkv[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sEkv[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sh_s[3]!="")
      {
      	//88	дБ, Высота сигнатуры
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sh_s[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sh_s[1]);
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sh_s[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sh_s[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sDelta_f_s[3]!="")
      {
      	//89	МГц, Ширина сигнатуры
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_f_s[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_f_s[1][0])+(unsigned char)(cDiD->sDelta_f_s[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
         oSel.SetText((cDiD->sDelta_f_s[1]).Right(2));
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_f_s[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_f_s[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sS_s[3]!="")
      {
      	//90	МГц, Площадь сигнатуры
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sS_s[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sS_s[1]);
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sS_s[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sS_s[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sI_ekv[3]!="")
      {
      	//91	дБ, Выигрыш за счет эквалайзера
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sI_ekv[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sI_ekv[1]);
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sI_ekv[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sI_ekv[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sV_min_eff_pred[3]!="")
      {
      	//92	дБ, Предельно реализуемый минимально допустимый эффективный множитель интерференционного ослабления без учета эквалайзера
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_eff_pred[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sV_min_eff_pred[1]);
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_eff_pred[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_eff_pred[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sV_min_eff_pred_ekv[3]!="")
      {
      	//93	дБ, Предельно реализуемый минимально допустимый эффективный множитель интерференционного ослабления с учетом эквалайзера
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_eff_pred_ekv[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sV_min_eff_pred_ekv[1]);
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_eff_pred_ekv[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_eff_pred_ekv[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sV_min_eff[3]!="")
      {
      	//94	дБ, Минимально допустимый эффективный множитель интерференционного ослабления
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_eff[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sV_min_eff[1]);
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_eff[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_eff[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sR_pr[3]!="")
      {
      	//95	км, Расстояние до критического препятствия по условию минимума относительного просвета при отсутствии рефракции
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_pr[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_pr[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_pr[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_pr[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sH_pr[3]!="")
      {
      	//96	м, Просвет в точке критического препятствия по условию минимума относительного просвета при отсутствии рефракции
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_pr[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_pr[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_pr[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_pr[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      //97	км, Расстояние до критического препятствия при отсутствии рефракции
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sR_kr[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sR_kr[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sR_kr[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sR_kr[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //98	Относительная координата критического препятствия при отсутствии рефракции
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sk[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sk[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sk[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sk[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //99	м, Просвет в точке критического препятствия при отсутствии рефракции
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sH_kr[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sH_kr[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sH_kr[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sH_kr[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //100	м, Оптимальный просвет в точке критического препятствия при отсутствии рефракции
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sH_0[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sH_0[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sH_0[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sH_0[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //101	км, Параметр хорды при отсутствии рефракции
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sParametr_r[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sParametr_r[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sParametr_r[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sParametr_r[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //102	Отношение длины хорды к протяженности интервала при отсутствии рефракции
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->slr[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->slr[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->slr[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->slr[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //103	м, Высота сегмента аппроксимирующей сферы при отсутствии рефракции
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sDelta_y[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_y[1][0])+(unsigned char)(cDiD->sDelta_y[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      oSel.SetText((cDiD->sDelta_y[1]).Right(5));
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sDelta_y[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sDelta_y[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //104	Отношение высоты аппроксимирующей сферы к оптимальному просвету при отсутствии рефракции
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sAlfa_delta_y[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.InsertSymbol((256*(unsigned char)(cDiD->sAlfa_delta_y[1][0])+(unsigned char)(cDiD->sAlfa_delta_y[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      oSel.InsertSymbol((256*(unsigned char)(cDiD->sAlfa_delta_y[1][2])+(unsigned char)(cDiD->sAlfa_delta_y[1][3])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      oSel.SetText((cDiD->sAlfa_delta_y[1]).Right(5));
      oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sAlfa_delta_y[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sAlfa_delta_y[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //105	Параметр, характеризующий кривизну аппроксимирующей сферы при отсутствии рефракции
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->smu_0[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.InsertSymbol((256*(unsigned char)(cDiD->smu_0[1][0])+(unsigned char)(cDiD->smu_0[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      oSel.SetText((cDiD->smu_0[1]).Right(1));
      oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->smu_0[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->smu_0[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //106	Параметр А для критического препятствия при отсутствии рефракции
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sA_kr[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sA_kr[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sA_kr[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sA_kr[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //107	км, Эквивалентный радиус земли при средней рефракции
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sa_e_km_2[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sa_e_km_2[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sa_e_km_2[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sa_e_km_2[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      if(cDiD->sR_pr_2[3]!="")
      {
      	//108	км, Расстояние до критического препятствия по условию минимума относительного просвета при средней рефракции
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_pr_2[0]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(64)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_pr_2[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_pr_2[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_pr_2[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sH_pr_2[3]!="")
      {
      	//109	м, Просвет в точке критического препятствия по условию минимума относительного просвета при средней рефракции
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_pr_2[0]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(63)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_pr_2[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_pr_2[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_pr_2[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      //110	км, Расстояние до критического препятствия при средней рефракции
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sR_kr_2[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sR_kr_2[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sR_kr_2[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sR_kr_2[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //111	Относительная координата критического препятствия при средней рефракции
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sk_2[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sk_2[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sk_2[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sk_2[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //112	м, Просвет в точке критического препятствия при средней рефракции
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sH_kr_2[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sH_kr_2[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sH_kr_2[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sH_kr_2[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //113	м, Оптимальный просвет в точке критического препятствия при средней рефракции
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sH_0_2[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sH_0_2[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sH_0_2[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sH_0_2[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //114	Относительный просвет в точке критического препятствия при средней рефракции
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sp_ot_g_sr_kr[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sp_ot_g_sr_kr[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(3)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sp_ot_g_sr_kr[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sp_ot_g_sr_kr[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //115	км, Параметр хорды при средней рефракции
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sParametr_r_2[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sParametr_r_2[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sParametr_r_2[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sParametr_r_2[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //116	м, Высота сегмента аппроксимирующей сферы при средней рефракции
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sDelta_y_2[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_y_2[1][0])+(unsigned char)(cDiD->sDelta_y_2[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      oSel.SetText((cDiD->sDelta_y_2[1]).Right(6));
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sDelta_y_2[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sDelta_y_2[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //117	kм, Радиус аппроксимирующей сферы при средней рефракции
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sb_2[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sb_2[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sb_2[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sb_2[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //118	Параметр, характеризующий кривизну аппроксимирующей сферы при средней рефракции
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->smu_2[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.InsertSymbol((256*(unsigned char)(cDiD->smu_2[1][0])+(unsigned char)(cDiD->smu_2[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      oSel.SetText((cDiD->smu_2[1]).Right(5));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->smu_2[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->smu_2[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //119	1/м, Пороговое значение эффективного градиента диэлектрической проницаемости воздуха при Vдифр=Vмин.
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sg_0[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sg_0[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sg_0[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sg_0[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //120	км, Эквивалентный радиус земли при пороговой рефракции
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sa_e_km_3[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sa_e_km_3[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sa_e_km_3[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sa_e_km_3[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //121	км, Расстояние до левой границы хорды при пороговой рефракции
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sR1x_3[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sR1x_3[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sR1x_3[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sR1x_3[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //122	км, Расстояние до правой границы хорды при пороговой рефракции
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sR2x_3[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sR2x_3[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sR2x_3[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sR2x_3[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //123	м, Высотная отметка левой границы хорды при пороговой рефракции
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sH1h_3[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sH1h_3[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sH1h_3[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sH1h_3[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //124	м, Высотная отметка правой границы хорды при пороговой рефракции
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sH2h_3[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sH2h_3[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sH2h_3[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sH2h_3[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //125	м, Высотная отметка хорды в точке наиболее высокого препятствия по критерию минимального относительного просвета при пороговой рефракции
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sH_kr_h_3[0]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(59)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sH_kr_h_3[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sH_kr_h_3[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sH_kr_h_3[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      if(cDiD->sR1_kas_3[3]!="")
      {
      	//126	км, Расстояние до точки касания левой касательной с поверхностью профиля при пороговой рефракции
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sR1_kas_3[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sR1_kas_3[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sR1_kas_3[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sR1_kas_3[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sR2_kas_3[3]!="")
      {
      	//127	км, Расстояние до точки касания правой касательной с поверхностью профиля при пороговой рефракции
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sR2_kas_3[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sR2_kas_3[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sR2_kas_3[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sR2_kas_3[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sH1_kas_3[3]!="")
      {
      	//128	м, Высотная отметка точки касания левой касательной с поверхностью профиля при пороговой рефракции
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_kas_3[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_kas_3[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_kas_3[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_kas_3[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sH2_kas_3[3]!="")
      {
      	//129	м, Высотная отметка точки касания правой касательной с поверхностью профиля при пороговой рефракции
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_kas_3[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_kas_3[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_kas_3[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_kas_3[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sR1_ellipse_3[3]!="")
      {
      	//130	км, Расстояние до точки пересечения левого полуэллипса с поверхностью профиля при пороговой рефракции
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sR1_ellipse_3[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sR1_ellipse_3[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sR1_ellipse_3[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sR1_ellipse_3[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sR2_ellipse_3[3]!="")
      {
      	//131	км, Расстояние до точки пересечения правого полуэллипса с поверхностью профиля при пороговой рефракции
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sR2_ellipse_3[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sR2_ellipse_3[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sR2_ellipse_3[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sR2_ellipse_3[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sH1_ellipse_3[3]!="")
      {
      	//132	м, Высотная отметка точки пересечения левого полуэллипса с поверхностью профиля при пороговой рефракции
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_ellipse_3[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_ellipse_3[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_ellipse_3[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_ellipse_3[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sH2_ellipse_3[3]!="")
      {
      	//133	м, Высотная отметка точки пересечения правого полуэллипса с поверхностью профиля при пороговой рефракции
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_ellipse_3[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_ellipse_3[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_ellipse_3[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_ellipse_3[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sR1_horda_H0_3[3]!="")
      {
      	//134	км, Расстояние до левой границы хорды, построенной по критерию H0, при пороговой рефракции
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sR1_horda_H0_3[0]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(51)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdExtend)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sR1_horda_H0_3[0][51])+(unsigned char)(cDiD->sR1_horda_H0_3[0][52])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(3)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sR1_horda_H0_3[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sR1_horda_H0_3[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sR1_horda_H0_3[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sR2_horda_H0_3[3]!="")
      {
      	//135	км, Расстояние до правой границы хорды, построенной по критерию H0, при пороговой рефракции
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sR2_horda_H0_3[0]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(52)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdExtend)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sR2_horda_H0_3[0][52])+(unsigned char)(cDiD->sR2_horda_H0_3[0][53])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(3)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sR2_horda_H0_3[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sR2_horda_H0_3[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sR2_horda_H0_3[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sH1_horda_H0_3[3]!="")
      {
      	//136	м, Высотная отметка левой границы хорды, построенной по критерию H0, при пороговой рефракции
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_horda_H0_3[0]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(51)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdExtend)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sH1_horda_H0_3[0][51])+(unsigned char)(cDiD->sH1_horda_H0_3[0][52])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(3)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_horda_H0_3[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_horda_H0_3[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_horda_H0_3[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sH2_horda_H0_3[3]!="")
      {
      	//137	м, Высотная отметка правой границы хорды, построенной по критерию H0, при пороговой рефракции
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_horda_H0_3[0]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(52)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdExtend)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sH2_horda_H0_3[0][52])+(unsigned char)(cDiD->sH2_horda_H0_3[0][53])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(3)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_horda_H0_3[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_horda_H0_3[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_horda_H0_3[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sR_cross_kas_3[3]!="")
      {
      	//138	км, Расстояние до точки пересечения касательных к поверхности профиля при пороговой рефракции
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_cross_kas_3[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_cross_kas_3[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_cross_kas_3[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_cross_kas_3[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sH_cross_kas_3[3]!="")
      {
      	//139	м, Высотная отметка точки пересечения касательных к поверхности профиля при пороговой рефракции
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_cross_kas_3[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_cross_kas_3[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_cross_kas_3[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_cross_kas_3[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sProsvet_cross_kas_3[3]!="")
      {
      	//140	м, Просвет в точке пересечения касательных при пороговой рефракции
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sProsvet_cross_kas_3[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sProsvet_cross_kas_3[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sProsvet_cross_kas_3[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sProsvet_cross_kas_3[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      //141	км, Расстояние до критического препятствия при пороговой рефракции
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sR_kr_3[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sR_kr_3[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sR_kr_3[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sR_kr_3[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //142	Относительная координата критического препятствия при пороговой рефракции
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sk_3[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sk_3[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sk_3[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sk_3[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //143	м, Просвет в точке критического препятствия при пороговой рефракции
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sH_kr_3[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sH_kr_3[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sH_kr_3[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sH_kr_3[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //144	м, Оптимальный просвет в точке критического препятствия при пороговой рефракции
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sH_0_3[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sH_0_3[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sH_0_3[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sH_0_3[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //145	Относительный просвет в точке критического препятствия при пороговой рефракции
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sp_ot_g_0_kr[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sp_ot_g_0_kr[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(3)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(3)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sp_ot_g_0_kr[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sp_ot_g_0_kr[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //146	км, Параметр хорды при пороговой рефракции
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sParametr_r_3[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sParametr_r_3[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sParametr_r_3[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sParametr_r_3[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //147	м, Высота сегмента аппроксимирующей сферы при пороговой рефракции
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sDelta_y_3[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_y_3[1][0])+(unsigned char)(cDiD->sDelta_y_3[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      oSel.SetText((cDiD->sDelta_y_3[1]).Right(5));
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sDelta_y_3[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sDelta_y_3[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //148	kм, Радиус аппроксимирующей сферы при пороговой рефракции
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sb_3[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sb_3[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sb_3[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sb_3[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //149	Параметр, характеризующий кривизну аппроксимирующей сферы при пороговой рефракции
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->smu_3[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.InsertSymbol((256*(unsigned char)(cDiD->smu_3[1][0])+(unsigned char)(cDiD->smu_3[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      oSel.SetText((cDiD->smu_3[1]).Right(4));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->smu_3[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->smu_3[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //150	дБ, Дифракционное ослабление при пороговой рефракции
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sV_difr_ot_g_0[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sV_difr_ot_g_0[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sV_difr_ot_g_0[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sV_difr_ot_g_0[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //151	Параметр пси
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sPsi[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.InsertSymbol((256*(unsigned char)(cDiD->sPsi[1][0])+(unsigned char)(cDiD->sPsi[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sPsi[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sPsi[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //152	%, Неустойчивость, обусловленная рефракционными явлениями
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sT_0[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sT_0[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sT_0[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sT_0[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      if(cDiD->sR_otr[3]!="")
      {
      	//153	км, Координата возможной точки отражения при отсутствии рефракции
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_otr[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_otr[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_otr[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_otr[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sH_otr[3]!="")
      {
      	//154	м, Просвет в возможной точке отражения при отсутствии рефракции
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_otr[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_otr[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_otr[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_otr[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sH_0_otr[3]!="")
      {
      	//155	м, Оптимальный просвет в возможной точке отражения при отсутствии рефракции
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_0_otr[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_0_otr[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_0_otr[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_0_otr[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sp_ot_g_otr[3]!="")
      {
      	//156	Относительный просвет в возможной точке отражения при отсутствии рефракции
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sp_ot_g_otr[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sp_ot_g_otr[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(4)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sp_ot_g_otr[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sp_ot_g_otr[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sX_prodoln[3]!="")
      {
      	//157	км, Протяженность отражающего участка профиля, требуемая для слабопересеченного характера трассы при отсутствии рефракции
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sX_prodoln[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sX_prodoln[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sX_prodoln[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sX_prodoln[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sDelta_h_max[3]!="")
      {
      	//158	м, Максимальное отклонение рельефа отражающего участка профиля от аппроксимирующей поверхности, при котором еще можно считать, что профиль слабопересеченный, при отсутствии рефракции
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_h_max[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_h_max[1][0])+(unsigned char)(cDiD->sDelta_h_max[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
         oSel.SetText((cDiD->sDelta_h_max[1]).Right(10));
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_h_max[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_h_max[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sR_otr_2[3]!="")
      {
      	//159	км, Расстояние до точки отражения для основных антенн при средней рефракции
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_otr_2[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_otr_2[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_otr_2[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_otr_2[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sH_otr_2[3]!="")
      {
      	//160	м, Просвет в точке отражения для основных антенн при средней рефракции
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_otr_2[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_otr_2[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_otr_2[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_otr_2[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sH_0_otr_2[3]!="")
      {
      	//161	м, Оптимальный просвет в точке отражения для основных антенн при средней рефракции
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_0_otr_2[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_0_otr_2[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_0_otr_2[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_0_otr_2[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sp_ot_g_sr_otr[3]!="")
      {
      	//162	Относительный просвет в точке отражения для основных антенн при средней рефракции
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sp_ot_g_sr_otr[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sp_ot_g_sr_otr[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(4)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(3)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sp_ot_g_sr_otr[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sp_ot_g_sr_otr[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sX_prodoln_2[3]!="")
      {
      	//163	км, Протяженность отражающего участка профиля, требуемая для слабопересеченного характера трассы, при средней рефракции
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sX_prodoln_2[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sX_prodoln_2[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sX_prodoln_2[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sX_prodoln_2[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sDelta_h_max_2[3]!="")
      {
      	//164	м, Максимальное отклонение рельефа отражающего участка профиля от аппроксимирующей поверхности, при котором еще можно считать, что профиль слабопересеченный, при средней рефракции
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_h_max_2[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_h_max_2[1][0])+(unsigned char)(cDiD->sDelta_h_max_2[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
         oSel.SetText((cDiD->sDelta_h_max_2[1]).Right(11));
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_h_max_2[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_h_max_2[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sA_otr[3]!="")
      {
      	//165	Параметр А для точки отражения при средней рефракции
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sA_otr[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sA_otr[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sA_otr[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sA_otr[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sF_ot_p_g_A[3]!="")
      {
      	//166	Параметр F(p(g),A) для точки отражения при средней рефракции
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sF_ot_p_g_A[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sF_ot_p_g_A[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(5)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sF_ot_p_g_A[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sF_ot_p_g_A[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sK_vp[3]!="")
      {
      	//167	%, Коэффициент водной поверхности
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_vp[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_vp[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_vp[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_vp[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      //168	Параметр Q при средней рефракции
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sQ[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sQ[1]);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sQ[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sQ[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //169	%, Параметр, учитывающий вероятность возникновения многолучевых замираний, обусловленных отражениями радиоволн от слоистых неоднородностей тропосферы с перепадом диэлектрической проницаемости воздуха Delta_epsilon
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sT_Delta_eps[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText((cDiD->sT_Delta_eps[1]).Left(2));
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdMove)));
      oSel.InsertSymbol((256*(unsigned char)(cDiD->sT_Delta_eps[1][2])+(unsigned char)(cDiD->sT_Delta_eps[1][3])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      oSel.InsertSymbol((256*(unsigned char)(cDiD->sT_Delta_eps[1][4])+(unsigned char)(cDiD->sT_Delta_eps[1][5])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      oSel.SetText((cDiD->sT_Delta_eps[1]).Right(1));
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sT_Delta_eps[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sT_Delta_eps[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      if(cDiD->ssigma_0[3]!="")
      {
      	//170	дБ, Стандартное отклонение логарифмически нормального закона распределения V в области T(V)>1%
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->ssigma_0[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->ssigma_0[1][0])+(unsigned char)(cDiD->ssigma_0[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->ssigma_0[1]).Right(1));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->ssigma_0[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->ssigma_0[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->ssigma_1[3]!="")
      {
      	//171	дБ, Стандартное отклонение логарифмически нормального закона распределения V в области T(V)<1%
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->ssigma_1[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->ssigma_1[1][0])+(unsigned char)(cDiD->ssigma_1[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->ssigma_1[1]).Right(1));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->ssigma_1[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->ssigma_1[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sT_int_priz[3]!="")
      {
      	//172.1	%, Неустойчивость, обусловленная интерференционными явлениями на приземных трассах
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_int_priz[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_int_priz[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_int_priz[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_int_priz[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }

      if(cDiD->sT_int_vg[3]!="")
      {
      	//172.2	%, Неустойчивость, обусловленная интерференционными явлениями на высокогорных трассах
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_int_vg[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_int_vg[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_int_vg[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_int_vg[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      //173	%, Неустойчивость, обусловленная интерференционными явлениями
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sT_int[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sT_int[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sT_int[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sT_int[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      if(cDiD->sRaion_d[3]!="")
      {
      	//174	Номер района по интенсивности дождей
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sRaion_d[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sRaion_d[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sRaion_d[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sRaion_d[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sRaion_Qd[3]!="")
      {
      	//175	Номер района распределения Qд
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sRaion_Qd[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sRaion_Qd[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sRaion_Qd[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sRaion_Qd[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sI_max[3]!="")
      {
      	//176	мм/ч, Максимально допустимая интенсивность дождя
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sI_max[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sI_max[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sI_max[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sI_max[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sT_d_m[3]!="")
      {
      	//177	%, Неустойчивость, обусловленная влиянием осадков в наихудший месяц
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_d_m[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_d_m[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_d_m[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_d_m[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sQ_d[3]!="")
      {
      	//178	Коэффициент пересчета месячной статистики дождей к годовой
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sQ_d[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sQ_d[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sQ_d[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sQ_d[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sT_d_g[3]!="")
      {
      	//179	%, Неустойчивость, обусловленная влиянием осадков в среднем за год
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_d_g[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_d_g[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_d_g[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_d_g[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sPsi_tau_0[3]!="")
      {
      	//180	км^2, Обобщенный параметр для определения C_m_0
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sPsi_tau_0[0]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(37)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sPsi_tau_0[1][0])+(unsigned char)(cDiD->sPsi_tau_0[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sPsi_tau_0[1][2])+(unsigned char)(cDiD->sPsi_tau_0[1][3])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sPsi_tau_0[1]).Right(5));
      	oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sPsi_tau_0[2]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSuperscript(True);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sPsi_tau_0[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sC_m_0[3]!="")
      {
      	//181	с, Эмпирический коэффициент для расчета tau_m_0
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_m_0[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_m_0[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_m_0[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_m_0[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->stau_m_0[3]!="")
      {
      	//182	с, Медианное значение длительности замираний в условиях субрефракционных замираний
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->stau_m_0[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->stau_m_0[1][0])+(unsigned char)(cDiD->stau_m_0[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->stau_m_0[1]).Right(7));
      	oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->stau_m_0[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->stau_m_0[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->ssigma_tau_0[3]!="")
      {
      	//183	дБ, Стандартное отклонение для логарифма длительности замираний в условиях субрефр. замираний
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->ssigma_tau_0[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->ssigma_tau_0[1][0])+(unsigned char)(cDiD->ssigma_tau_0[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->ssigma_tau_0[1][2])+(unsigned char)(cDiD->ssigma_tau_0[1][3])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->ssigma_tau_0[1]).Right(5));
      	oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->ssigma_tau_0[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->ssigma_tau_0[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sPsi_tau_int[3]!="")
      {
      	//184	км^2, Обобщенный параметр для определения C_m_int
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sPsi_tau_int[0]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(37)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sPsi_tau_int[1][0])+(unsigned char)(cDiD->sPsi_tau_int[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sPsi_tau_int[1][2])+(unsigned char)(cDiD->sPsi_tau_int[1][3])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sPsi_tau_int[1]).Right(4));
      	oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sPsi_tau_int[2]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSuperscript(True);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sPsi_tau_int[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sC_m_int[3]!="")
      {
      	//185	с, Эмпирический коэффициент для расчета tau_m_int
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_m_int[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_m_int[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_m_int[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_m_int[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->stau_m_int[3]!="")
      {
      	//186	с, Медианное значение длительности замираний в условиях интерференционных замираний
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->stau_m_int[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->stau_m_int[1][0])+(unsigned char)(cDiD->stau_m_int[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->stau_m_int[1]).Right(6));
      	oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->stau_m_int[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->stau_m_int[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->ssigma_tau_int[3]!="")
      {
      	//187	дБ, Стандартное отклонение для логарифма длительности замираний в условиях интерф. замираний
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->ssigma_tau_int[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->ssigma_tau_int[1][0])+(unsigned char)(cDiD->ssigma_tau_0[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->ssigma_tau_int[1][2])+(unsigned char)(cDiD->ssigma_tau_0[1][3])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->ssigma_tau_int[1]).Right(4));
      	oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->ssigma_tau_int[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->ssigma_tau_int[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sfi_tau_0[3]!="")
      {
      	//188	Коэффициент готовности в условиях субрефракционных замираний
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sfi_tau_0[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sfi_tau_0[1][0])+(unsigned char)(cDiD->sfi_tau_0[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sfi_tau_0[1][2])+(unsigned char)(cDiD->sfi_tau_0[1][3])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sfi_tau_0[1]).Right(1));
      	oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sfi_tau_0[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sfi_tau_0[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sfi_tau_int[3]!="")
      {
      	//189	Коэффициент готовности в условиях интерференционных замираний
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sfi_tau_int[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sfi_tau_int[1][0])+(unsigned char)(cDiD->sfi_tau_int[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sfi_tau_int[1][2])+(unsigned char)(cDiD->sfi_tau_int[1][3])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sfi_tau_int[1]).Right(4));
      	oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sfi_tau_int[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sfi_tau_int[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      //190	Коэффициент интерференции
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sK_int[0]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(66)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(6)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sK_int[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sK_int[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sK_int[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //191	%, Коэффициент секунд со значительным количеством ошибок при одинарном приеме
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sSESR[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sSESR[1]);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sSESR[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sSESR[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //192	%, Коэффициент неготовности при одинарном приеме
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sK_ng[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sK_ng[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sK_ng[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sK_ng[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //193	Учет частотно-разнесенного приема (есть/нет)
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sChRP[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sChRP[1]);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sChRP[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sChRP[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      if(cDiD->sK_stv[3]!="")
      {
      	//194	Количество рабочих стволов, не учитывая резервного
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_stv[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_stv[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_stv[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_stv[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sDelta_f[3]!="")
      {
      	//195	МГц, Разнос по частоте между резервным стволом и ближайшем к нему рабочим
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_f[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_f[1][0])+(unsigned char)(cDiD->sDelta_f[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sDelta_f[1]).Right(1));
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_f[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_f[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sDelta_f_f[3]!="")
      {
      	//196	%, Отношение частотного разноса к рабочей частоте
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_f_f[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_f_f[1][0])+(unsigned char)(cDiD->sDelta_f_f[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sDelta_f_f[1]).Right(3));
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_f_f[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_f_f[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sC_Delta_f_priz[3]!="")
      {
      	//197	Эмпирический коэффициент, учитывающий статистическую зависимость замираний на интервале РРЛ при частотном разнесении двух высокочастотных стволов на величину Delta_f, а также особенности работы системы резервирования при интерференционных замираниях для приземных трасс
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_Delta_f_priz[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText((cDiD->sC_Delta_f_priz[1]).Left(1));
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sC_Delta_f_priz[1][1])+(unsigned char)(cDiD->sC_Delta_f_priz[1][2])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sC_Delta_f_priz[1]).Right(3));
      	oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_Delta_f_priz[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_Delta_f_priz[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sC_f_priz[3]!="")
      {
      	//198	Коэффициент C_f для приземных трасс
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_f_priz[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_f_priz[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_f_priz[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_f_priz[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sC_Delta_f[3]!="")
      {
      	//199	Эмпирический коэффициент, учитывающий статистическую зависимость замираний на интервале РРЛ при частотном разнесении двух высокочастотных стволов на величину Delta_f, а также особенности работы системы резервирования при интерференционных замираниях
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_Delta_f[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText((cDiD->sC_Delta_f[1]).Left(1));
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sC_Delta_f[1][1])+(unsigned char)(cDiD->sC_Delta_f[1][2])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sC_Delta_f[1]).Right(1));
      	oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_Delta_f[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_Delta_f[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sC_f[3]!="")
      {
      	//200	Коэффициент C_f
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_f[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_f[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_f[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_f[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sq[3]!="")
      {
      	//201	Коэффициент, учитывающий часть времени, в течение которого ствол горячего резерва не используется для резервирования при замираниях
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sq[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sq[1]);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sq[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sq[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sT_int_ChRP_priz[3]!="")
      {
      	//202	%, Интерференционная неустойчивость на интервале ЦРРЛ с ЧРП в худший месяц для приземных трасс
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_int_ChRP_priz[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_int_ChRP_priz[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_int_ChRP_priz[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_int_ChRP_priz[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->salfa_tau_int_ChRP[3]!="")
      {
      	//203	Коэффициент неготовности системы с ЧРП в условиях интерференционных замираний
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->salfa_tau_int_ChRP[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.InsertSymbol((256*(unsigned char)(cDiD->salfa_tau_int_ChRP[1][0])+(unsigned char)(cDiD->salfa_tau_int_ChRP[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->salfa_tau_int_ChRP[1][2])+(unsigned char)(cDiD->salfa_tau_int_ChRP[1][3])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->salfa_tau_int_ChRP[1]).Right(4));
      	oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->salfa_tau_int_ChRP[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->salfa_tau_int_ChRP[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sK_ng_ChRP_mes[3]!="")
      {
      	//204	%, Коэффициент неготовности на интервале ЦРРЛ с ЧРП в худший месяц
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_ChRP_mes[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_ChRP_mes[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_ChRP_mes[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_ChRP_mes[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sT_ChRP[3]!="")
      {
      	//205	%, Суммарная неустойчивость связи на интервале ЦРРЛ с ЧРП в худший месяц
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_ChRP[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_ChRP[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_ChRP[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_ChRP[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sSESR_ChRP[3]!="")
      {
      	//206	%, Коэффициент секунд со значительным количеством ошибок при ЧРП на интервале
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sSESR_ChRP[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sSESR_ChRP[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(4)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sSESR_ChRP[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sSESR_ChRP[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sK_ng_ChRP[3]!="")
      {
      	//207	%, Коэффициент неготовности на интервале ЦРРЛ с ЧРП в среднем за год
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_ChRP[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_ChRP[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_ChRP[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_ChRP[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sPRP[3]!="")
      {
      	//208	Учет пространственно-разнесенного приема (есть/нет)
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sPRP[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sPRP[1]);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sPRP[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sPRP[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sDelta_h_rek_sleva[3]!="")
      {
      	//209	м, Рекомендуемый разнос антенн при ПРП слева
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_h_rek_sleva[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_h_rek_sleva[1][0])+(unsigned char)(cDiD->sDelta_h_rek_sleva[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sDelta_h_rek_sleva[1]).Right(10));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(9)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_h_rek_sleva[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_h_rek_sleva[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sDelta_h_rek_sprava[3]!="")
      {
      	//210	м, Рекомендуемый разнос антенн при ПРП справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_h_rek_sprava[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_h_rek_sprava[1][0])+(unsigned char)(cDiD->sDelta_h_rek_sprava[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sDelta_h_rek_sprava[1]).Right(11));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(10)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_h_rek_sprava[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_h_rek_sprava[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sShema_PRP[3]!="")
      {
      	//211	Схема пространственно-разнесенного приема ("классическая (осн.-доп.)"/"крест-на-крест (доп.-доп.)")
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sShema_PRP[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sShema_PRP[1]);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sShema_PRP[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sShema_PRP[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sL_AVT2_dop[3]!="")
      {
      	//218	м, Длина волновода дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sL_AVT2_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sL_AVT2_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sL_AVT2_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sL_AVT2_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sh_sr_dop[3]!="")
      {
      	//219	м, Средняя высота трассы луча над уровнем моря для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sh_sr_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sh_sr_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sh_sr_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sh_sr_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sGorn_dop[3]!="")
      {
      	//220	Горная трасса для дополнительной антенны справа (приземная/горная/высокогорная)
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sGorn_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sGorn_dop[1]);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sGorn_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sGorn_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sEta_AFT_dop[3]!="")
      {
      	//221	дБ, Суммарные потери на интервале в антенно-фидерных трактах для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sEta_AFT_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sEta_AFT_dop[1][0])+(unsigned char)(cDiD->sEta_AFT_dop[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sEta_AFT_dop[1]).Right(7));
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sEta_AFT_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sEta_AFT_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sW_0_dop[3]!="")
      {
      	//222	дБ, Ослабление сигнала в свободном пространстве для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sW_0_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sW_0_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sW_0_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sW_0_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sR_pr_dop[3]!="")
      {
      	//223	км, Расстояние до критического препятствия по условию минимума относительного просвета при средней рефракции для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_pr_dop[0]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(54)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_pr_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_pr_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_pr_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sH_pr_dop[3]!="")
      {
      	//224	м, Просвет в точке критического препятствия по условию минимума относительного просвета при средней рефракции для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_pr_dop[0]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(56)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_pr_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_pr_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_pr_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sR1x_dop[3]!="")
      {
      	//225	км, Расстояние до левой границы хорды при средней рефракции для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sR1x_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sR1x_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sR1x_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sR1x_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sR2x_dop[3]!="")
      {
      	//226	км, Расстояние до правой границы хорды при средней рефракции для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sR2x_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sR2x_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sR2x_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sR2x_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sH1h_dop[3]!="")
      {
      	//227	м, Высотная отметка левой границы хорды при средней рефракции для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1h_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1h_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1h_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1h_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sH2h_dop[3]!="")
      {
      	//228	м, Высотная отметка правой границы хорды при средней рефракции для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2h_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2h_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2h_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2h_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sH_kr_h_dop[3]!="")
      {
      	//229	м, Высотная отметка хорды в точке наиболее высокого препятствия по критерию минимального относительного просвета при средней рефракции для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_kr_h_dop[0]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(52)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_kr_h_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_kr_h_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_kr_h_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sR1_kas_dop[3]!="")
      {
      	//230	км, Расстояние до точки касания левой касательной с поверхностью профиля при средней рефракции для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sR1_kas_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sR1_kas_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sR1_kas_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sR1_kas_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sR2_kas_dop[3]!="")
      {
      	//231	км, Расстояние до точки касания правой касательной с поверхностью профиля при средней рефракции для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sR2_kas_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sR2_kas_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sR2_kas_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sR2_kas_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sH1_kas_dop[3]!="")
      {
      	//232	м, Высотная отметка точки касания левой касательной с поверхностью профиля при средней рефракции для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_kas_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_kas_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_kas_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_kas_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sH2_kas_dop[3]!="")
      {
      	//233	м, Высотная отметка точки касания правой касательной с поверхностью профиля при средней рефракции для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_kas_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_kas_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_kas_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_kas_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sR1_ellipse_dop[3]!="")
      {
      	//234	км, Расстояние до точки пересечения левого полуэллипса с поверхностью профиля при средней рефракции для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sR1_ellipse_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sR1_ellipse_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sR1_ellipse_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sR1_ellipse_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sR2_ellipse_dop[3]!="")
      {
      	//235	км, Расстояние до точки пересечения правого полуэллипса с поверхностью профиля при средней рефракции для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sR2_ellipse_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sR2_ellipse_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sR2_ellipse_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sR2_ellipse_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sH1_ellipse_dop[3]!="")
      {
      	//236	м, Высотная отметка точки пересечения левого полуэллипса с поверхностью профиля при средней рефракции для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_ellipse_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_ellipse_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_ellipse_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_ellipse_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sH2_ellipse_dop[3]!="")
      {
      	//237	м, Высотная отметка точки пересечения правого полуэллипса с поверхностью профиля при средней рефракции для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_ellipse_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_ellipse_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_ellipse_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_ellipse_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sR1_horda_H0_dop[3]!="")
      {
      	//238	км, Расстояние до левой границы хорды, построенной по критерию H0, при средней рефракции для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sR1_horda_H0_dop[0]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(41)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdExtend)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sR1_horda_H0_dop[0][41])+(unsigned char)(cDiD->sR1_horda_H0_dop[0][42])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(3)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sR1_horda_H0_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sR1_horda_H0_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sR1_horda_H0_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sR2_horda_H0_dop[3]!="")
      {
      	//239	км, Расстояние до правой границы хорды, построенной по критерию H0, при средней рефракции для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sR2_horda_H0_dop[0]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(42)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdExtend)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sR2_horda_H0_dop[0][42])+(unsigned char)(cDiD->sR2_horda_H0_dop[0][43])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(3)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sR2_horda_H0_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sR2_horda_H0_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sR2_horda_H0_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sH1_horda_H0_dop[3]!="")
      {
      	//240	м, Высотная отметка левой границы хорды, построенной по критерию H0, при средней рефракции для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_horda_H0_dop[0]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(41)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdExtend)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sH1_horda_H0_dop[0][41])+(unsigned char)(cDiD->sH1_horda_H0_dop[0][42])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(3)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_horda_H0_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_horda_H0_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_horda_H0_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sH2_horda_H0_dop[3]!="")
      {
      	//241	м, Высотная отметка правой границы хорды, построенной по критерию H0, при средней рефракции для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_horda_H0_dop[0]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(42)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdExtend)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sH2_horda_H0_dop[0][42])+(unsigned char)(cDiD->sH2_horda_H0_dop[0][43])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(3)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_horda_H0_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_horda_H0_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_horda_H0_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sR_cross_kas_dop[3]!="")
      {
      	//242	км, Расстояние до точки пересечения касательных к поверхности профиля при средней рефракции для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_cross_kas_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_cross_kas_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_cross_kas_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_cross_kas_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sH_cross_kas_dop[3]!="")
      {
      	//243	м, Высотная отметка точки пересечения касательных к поверхности профиля при средней рефракции для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_cross_kas_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_cross_kas_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_cross_kas_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_cross_kas_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sProsvet_cross_kas_dop[3]!="")
      {
      	//244	м, Просвет в точке пересечения касательных при средней рефракции для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sProsvet_cross_kas_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sProsvet_cross_kas_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sProsvet_cross_kas_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sProsvet_cross_kas_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sR_kr_dop[3]!="")
      {
      	//245	км, Расстояние до критического препятствия при средней рефракции для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_kr_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_kr_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_kr_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_kr_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sk_dop[3]!="")
      {
      	//246	Относительная координата критического препятствия при средней рефракции для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sk_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sk_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sk_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sk_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sH_kr_dop[3]!="")
      {
      	//247	м, Просвет в точке критического препятствия при средней рефракции для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_kr_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_kr_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_kr_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_kr_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sH_0_dop[3]!="")
      {
      	//248	м, Оптимальный просвет в точке критического препятствия при средней рефракции для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_0_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_0_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_0_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_0_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sp_ot_g_kr_dop[3]!="")
      {
      	//249	Относительный просвет в точке критического препятствия при средней рефракции для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sp_ot_g_kr_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sp_ot_g_kr_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(7)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(3)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sp_ot_g_kr_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sp_ot_g_kr_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sParametr_r_dop[3]!="")
      {
      	//250	км, Параметр хорды при средней рефракции для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sParametr_r_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sParametr_r_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sParametr_r_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sParametr_r_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sDelta_y_dop[3]!="")
      {
      	//251	м, Высота сегмента аппроксимирующей сферы при средней рефракции для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_y_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_y_dop[1][0])+(unsigned char)(cDiD->sDelta_y_dop[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sDelta_y_dop[1]).Right(8));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(7)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_y_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_y_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sb_dop[3]!="")
      {
      	//252	kм, Радиус аппроксимирующей сферы при средней рефракции для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sb_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sb_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sb_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sb_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->smu_dop[3]!="")
      {
      	//253	Параметр, характеризующий кривизну аппроксимирующей сферы при средней рефракции для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->smu_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->smu_dop[1][0])+(unsigned char)(cDiD->smu_dop[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->smu_dop[1]).Right(7));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->smu_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->smu_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sV_difr_sr_dop[3]!="")
      {
      	//254	дБ, Дифракционное ослабление при средней рефракции для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_difr_sr_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_difr_sr_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_difr_sr_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_difr_sr_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sV_50_proc_dop[3]!="")
      {
      	//255	дБ, Среднее ослабление сигнала из-за перепада высот на горных и высокогорных трассах для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_50_proc_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_50_proc_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(4)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_50_proc_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_50_proc_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sP_pm_dop[3]!="")
      {
      	//256	дБм, Средний уровень сигнала на входе приемника для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sP_pm_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sP_pm_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sP_pm_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sP_pm_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sV_0_min_d_dop[3]!="")
      {
      	//257	дБ, Минимально допустимый множитель ослабления по дождям без учета деградации порогового уровня, средних ослаблений, с учетом поправки на типовые параметры оборудования для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_0_min_d_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_0_min_d_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_0_min_d_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_0_min_d_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sV_0_min_0_dop[3]!="")
      {
      	//258	дБ, Минимально допустимый множитель ослабления по субрефракции и инт. без учета деградации порогового уровня, средних ослаблений, с учетом поправки на типовые параметры оборудования для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_0_min_0_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_0_min_0_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_0_min_0_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_0_min_0_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sZ_por_dop_dop[3]!="")
      {
      	//259	дБ, Пороговое отношение (помеха/Рпм пор, при деградации порогового уровня на 3 дБ) для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_por_dop_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_por_dop_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_por_dop_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_por_dop_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sZ_sosedn_kanal_dop[3]!="")
      {
      	//260	дБ, Превышение отношением (помеха/Рпм пор, при деградации порогового уровня на 3 дБ) порогового отношения (помехи от соседнего канала/Рпм пор) для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_sosedn_kanal_dop[0]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(9)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(8)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(8)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(10)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_sosedn_kanal_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_sosedn_kanal_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_sosedn_kanal_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sDelta_V_sosedn_kanal_dop[3]!="")
      {
      	//261	дБ, Деградация порогового уровня из-за влияния помех от соседнего канала для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_sosedn_kanal_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_V_sosedn_kanal_dop[1][0])+(unsigned char)(cDiD->sDelta_V_sosedn_kanal_dop[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sDelta_V_sosedn_kanal_dop[1]).Right(7));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(6)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_sosedn_kanal_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_sosedn_kanal_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sD_p_0_dop[3]!="")
      {
      	//262	дБ, Коэффициент поляризационной защиты при отсутствии замираний дополнительной приемной антенны
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sD_p_0_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sD_p_0_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sD_p_0_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sD_p_0_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sQ_p_cochannel_dop[3]!="")
      {
      	//263	дБ, Коэффициент, учитывающий наклон кроссполяризационной диаграммы направленности антенны для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sQ_p_cochannel_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sQ_p_cochannel_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sQ_p_cochannel_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sQ_p_cochannel_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sD_p_dop[3]!="")
      {
      	//264	дБ, Коэффициент поляризационной защиты в условиях интерференционных замираний для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sD_p_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sD_p_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sD_p_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sD_p_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sD_p_dojd_dop[3]!="")
      {
      	//265	дБ, Коэффициент поляризационной защиты, обусловленный влиянием дождей для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sD_p_dojd_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sD_p_dojd_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sD_p_dojd_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sD_p_dojd_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sZ_polyariz_cochannel_dop[3]!="")
      {
      	//266	дБ, Превышение отношением (кроссполяризационной помехи/Рпм пор) порогового отношения (помеха/Рпм пор, при деградации порогового уровня на 3 дБ), для учета в расчетах Тинт (в расчетах Т0 кроссполяризационная помеха не учитывается) для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_polyariz_cochannel_dop[0]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(9)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(8)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(7)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(12)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(4)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_polyariz_cochannel_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_polyariz_cochannel_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_polyariz_cochannel_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sZ_polyariz_cochannel_dojd_dop[3]!="")
      {
      	//267	дБ, Превышение отношением (кроссполяризационной помехи/Рпм пор) порогового отношения (помеха/Рпм пор, при деградации порогового уровня на 3 дБ), для учета в расчетах Тд (в расчетах Т0 кроссполяризационная помеха не учитывается) для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_polyariz_cochannel_dojd_dop[0]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(9)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(8)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(7)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(12)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(4)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_polyariz_cochannel_dojd_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_polyariz_cochannel_dojd_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_polyariz_cochannel_dojd_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sDelta_V_polyariz_cochannel_dop[3]!="")
      {
      	//268	дБ, Деградация порогового уровня из-за влияния кроссполяризационных помех для учета в расчетах Тинт (в расчетах Т0 кроссполяризационная помеха не учитывается) для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_polyariz_cochannel_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_V_polyariz_cochannel_dop[1][0])+(unsigned char)(cDiD->sDelta_V_polyariz_cochannel_dop[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sDelta_V_polyariz_cochannel_dop[1]).Right(11));
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_polyariz_cochannel_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_polyariz_cochannel_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sDelta_V_polyariz_cochannel_dojd_dop[3]!="")
      {
      	//269	дБ, Деградация порогового уровня из-за влияния кроссполяризационных помех для учета в расчетах Тд (в расчетах Т0 кроссполяризационная помеха не учитывается) для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_polyariz_cochannel_dojd_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_V_polyariz_cochannel_dojd_dop[1][0])+(unsigned char)(cDiD->sDelta_V_polyariz_cochannel_dojd_dop[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sDelta_V_polyariz_cochannel_dojd_dop[1]).Right(9));
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_polyariz_cochannel_dojd_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_polyariz_cochannel_dojd_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sZ_por_dopoln[3]!="")
      {
      	//270	дБ, Отношение (суммарной помехи обратного направления и узлообразования/Рпм пор) для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_por_dopoln[0]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(49)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sZ_por_dopoln[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_por_dopoln[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_por_dopoln[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sZ_obr_uzl_dop[3]!="")
      {
      	//271	дБ, Превышение отношением (суммарной помехи обратного направления и узлообразования/Рпм пор) порогового отношения (помеха/Рпм пор, при деградации порогового уровня на 3 дБ) для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_obr_uzl_dop[0]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(12)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(11)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(16)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sZ_obr_uzl_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_obr_uzl_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_obr_uzl_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sDelta_V_obr_uzl_dop[3]!="")
      {
      	//272	дБ, Деградация порогового уровня из-за влияния помех с обратного направления и узлообразования для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_obr_uzl_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_V_obr_uzl_dop[1][0])+(unsigned char)(cDiD->sDelta_V_obr_uzl_dop[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
         oSel.SetText((cDiD->sDelta_V_obr_uzl_dop[1]).Right(13));
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_obr_uzl_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_obr_uzl_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sDelta_V_degr_int_dop[3]!="")
      {
      	//273	дБ, Деградация порогового уровня из-за влияния помех на Tинт для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_degr_int_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_V_degr_int_dop[1][0])+(unsigned char)(cDiD->sDelta_V_degr_int_dop[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
         oSel.SetText((cDiD->sDelta_V_degr_int_dop[1]).Right(14));
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_degr_int_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_degr_int_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sDelta_V_degr_subrefr_dop[3]!="")
      {
      	//274	дБ, Деградация порогового уровня из-за влияния помех на T0 для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_degr_subrefr_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_V_degr_subrefr_dop[1][0])+(unsigned char)(cDiD->sDelta_V_degr_subrefr_dop[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
         oSel.SetText((cDiD->sDelta_V_degr_subrefr_dop[1]).Right(15));
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_degr_subrefr_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_degr_subrefr_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sDelta_V_degr_dojd_dop[3]!="")
      {
      	//275	дБ, Деградация порогового уровня из-за влияния помех на Tд для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_degr_dojd_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_V_degr_dojd_dop[1][0])+(unsigned char)(cDiD->sDelta_V_degr_dojd_dop[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
         oSel.SetText((cDiD->sDelta_V_degr_dojd_dop[1]).Right(15));
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_degr_dojd_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_degr_dojd_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sV_min_int_dop[3]!="")
      {
      	//276	дБ, Минимально допустимый множитель интерференционного ослабления c учетом деградации порогового уровня, средних ослаблений и поправки на типовые параметры оборудования для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_int_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sV_min_int_dop[1]);
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_int_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_int_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sV_min_subrefr_dop[3]!="")
      {
      	//277	дБ, Минимально допустимый множитель субрефракционного ослабления c учетом деградации порогового уровня, средних ослаблений и поправки на типовые параметры оборудования для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_subrefr_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sV_min_subrefr_dop[1]);
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_subrefr_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_subrefr_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sV_min_eff_dop[3]!="")
      {
      	//278	дБ, Минимально допустимый эффективный множитель интерференционного ослабления для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_eff_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sV_min_eff_dop[1]);
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_eff_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_eff_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sR_otr_dop[3]!="")
      {
      	//279	км, Расстояние до точки отражения для дополнительной антенны справа при средней рефракции
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_otr_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sR_otr_dop[1]);
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_otr_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_otr_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sH_otr_dop[3]!="")
      {
      	//280	м, Просвет в точке отражения для дополнительной антенны справа при средней рефракции
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_otr_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sH_otr_dop[1]);
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_otr_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_otr_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sH_0_otr_dop[3]!="")
      {
      	//281	м, Оптимальный просвет в точке отражения для дополнительной антенны справа при средней рефракции
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_0_otr_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sH_0_otr_dop[1]);
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_0_otr_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_0_otr_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sp_ot_g_sr_otr_dop[3]!="")
      {
      	//282	Относительный просвет в точке отражения для дополнительной антенны справа при средней рефракции
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sp_ot_g_sr_otr_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sp_ot_g_sr_otr_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(8)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(3)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sp_ot_g_sr_otr_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sp_ot_g_sr_otr_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sX_prodoln_dop[3]!="")
      {
      	//283	км, Протяженность отражающего участка профиля, требуемая для слабопересеченного характера трассы, для дополнительной антенны справа при средней рефракции
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sX_prodoln_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sX_prodoln_dop[1]);
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sX_prodoln_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sX_prodoln_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sDelta_h_max_dop[3]!="")
      {
      	//284	м, Максимальное отклонение рельефа отражающего участка профиля от аппроксимирующей поверхности, при котором еще можно считать, что профиль слабопересеченный, дополнительной антенны справа при средней рефракции
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_h_max_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_h_max_dop[1][0])+(unsigned char)(cDiD->sDelta_h_max_dop[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
         oSel.SetText((cDiD->sDelta_h_max_dop[1]).Right(15));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(14)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_h_max_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_h_max_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sA_otr_dop[3]!="")
      {
      	//285	Параметр А для точки отражения при средней рефракции для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sA_otr_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sA_otr_dop[1]);
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sA_otr_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sA_otr_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sF_ot_p_g_A_dop[3]!="")
      {
      	//286	Параметр F(p(g),A) для точки отражения при средней рефракции для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sF_ot_p_g_A_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sF_ot_p_g_A_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(4)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(5)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sF_ot_p_g_A_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sF_ot_p_g_A_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sQ_dop[3]!="")
      {
      	//287	Параметр Q при средней рефракции для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sQ_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sQ_dop[1]);
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sQ_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sQ_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sT_int_dop[3]!="")
      {
      	//288	%, Неустойчивость, обусловленная интерференционными явлениями для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_int_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sT_int_dop[1]);
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_int_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_int_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sfi_tau_int_dop[3]!="")
      {
      	//289	Коэффициент готовности в условиях интерференционных замираний для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sfi_tau_int_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sfi_tau_int_dop[1][0])+(unsigned char)(cDiD->sfi_tau_int_dop[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sfi_tau_int_dop[1][2])+(unsigned char)(cDiD->sfi_tau_int_dop[1][3])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sfi_tau_int_dop[1]).Right(8));
      	oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sfi_tau_int_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sfi_tau_int_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sfi_tau_0_dop[3]!="")
      {
      	//290	Коэффициент готовности в условиях субрефракционных замираний для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sfi_tau_0_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sfi_tau_0_dop[1][0])+(unsigned char)(cDiD->sfi_tau_0_dop[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sfi_tau_0_dop[1][2])+(unsigned char)(cDiD->sfi_tau_0_dop[1][3])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sfi_tau_0_dop[1]).Right(5));
      	oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sfi_tau_0_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sfi_tau_0_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sK_int_dop[3]!="")
      {
      	//291	Коэффициент интерференции для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_int_dop[0]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(54)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(6)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_int_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_int_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_int_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sSESR_dop[3]!="")
      {
      	//292	%, Коэффициент секунд со значительным количеством ошибок при одинарном приеме для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sSESR_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sSESR_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(4)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sSESR_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sSESR_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sK_ng_dop[3]!="")
      {
      	//293	%, Коэффициент неготовности при одинарном приеме для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sPsi_tau_0_dop[3]!="")
      {
      	//294	км^2, Обобщенный параметр для определения C_m_0 для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sPsi_tau_0_dop[0]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(28)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sPsi_tau_0_dop[1][0])+(unsigned char)(cDiD->sPsi_tau_0_dop[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sPsi_tau_0_dop[1][2])+(unsigned char)(cDiD->sPsi_tau_0_dop[1][3])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sPsi_tau_0_dop[1]).Right(9));
      	oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sPsi_tau_0_dop[2]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSuperscript(True);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sPsi_tau_0_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sC_m_0_dop[3]!="")
      {
      	//295	с, Эмпирический коэффициент для расчета tau_m_0 для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_m_0_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_m_0_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_m_0_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_m_0_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->stau_m_0_dop[3]!="")
      {
      	//296	с, Медианное значение длительности замираний в условиях субрефракционных замираний для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->stau_m_0_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->stau_m_0_dop[1][0])+(unsigned char)(cDiD->stau_m_0_dop[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->stau_m_0_dop[1]).Right(11));
      	oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->stau_m_0_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->stau_m_0_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->ssigma_tau_0_dop[3]!="")
      {
      	//297	дБ, Стандартное отклонение для логарифма длительности замираний в условиях субрефр. замираний для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->ssigma_tau_0_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->ssigma_tau_0_dop[1][0])+(unsigned char)(cDiD->ssigma_tau_0_dop[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->ssigma_tau_0_dop[1][2])+(unsigned char)(cDiD->ssigma_tau_0_dop[1][3])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->ssigma_tau_0_dop[1]).Right(9));
      	oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->ssigma_tau_0_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->ssigma_tau_0_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sPsi_tau_int[3]!="")
      {
      	//298	км^2, Обобщенный параметр для определения C_m_int для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sPsi_tau_int_dop[0]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(33)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sPsi_tau_int_dop[1][0])+(unsigned char)(cDiD->sPsi_tau_int_dop[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sPsi_tau_int_dop[1][2])+(unsigned char)(cDiD->sPsi_tau_int_dop[1][3])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sPsi_tau_int_dop[1]).Right(8));
      	oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sPsi_tau_int_dop[2]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSuperscript(True);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sPsi_tau_int_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sC_m_int_dop[3]!="")
      {
      	//299	с, Эмпирический коэффициент для расчета tau_m_int для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_m_int_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_m_int_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_m_int_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_m_int_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->stau_m_int_dop[3]!="")
      {
      	//300	с, Медианное значение длительности замираний в условиях интерференционных замираний для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->stau_m_int_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->stau_m_int_dop[1][0])+(unsigned char)(cDiD->stau_m_int_dop[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->stau_m_int_dop[1]).Right(10));
      	oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->stau_m_int_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->stau_m_int_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->ssigma_tau_int_dop[3]!="")
      {
      	//301	дБ, Стандартное отклонение для логарифма длительности замираний в условиях интерф. замираний для дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->ssigma_tau_int_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->ssigma_tau_int_dop[1][0])+(unsigned char)(cDiD->ssigma_tau_int_dop[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->ssigma_tau_int_dop[1][2])+(unsigned char)(cDiD->ssigma_tau_int_dop[1][3])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->ssigma_tau_int_dop[1]).Right(8));
      	oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->ssigma_tau_int_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->ssigma_tau_int_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sC_Delta_h[3]!="")
      {
      	//302	Эмпирический коэффициент, учитывающий статистическую зависимость замираний при пространственном разнесении антенн (для основной антенны, если трасса слабопересеченная)
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_Delta_h[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText((cDiD->sC_Delta_h[1]).Left(1));
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sC_Delta_h[1][1])+(unsigned char)(cDiD->sC_Delta_h[1][2])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sC_Delta_h[1]).Right(5));
      	oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_Delta_h[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_Delta_h[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sC_Delta_h_dop[3]!="")
      {
      	//303	Эмпирический коэффициент, учитывающий статистическую зависимость замираний при пространственном разнесении антенн для дополнительной антенны для слабопересеченной трассы
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_Delta_h_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText((cDiD->sC_Delta_h_dop[1]).Left(1));
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sC_Delta_h_dop[1][1])+(unsigned char)(cDiD->sC_Delta_h_dop[1][2])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sC_Delta_h_dop[1]).Right(5));
      	oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_Delta_h_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_Delta_h_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sI_PRP_peres[3]!="")
      {
      	//304	дБ, Эффективность ПРП или выигрыш при ПРП по отношению к одинарному приему для пересеченной трассы
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sI_PRP_peres[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sI_PRP_peres[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sI_PRP_peres[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sI_PRP_peres[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sT_int_PRP[3]!="")
      {
      	//305	%, Неустойчивость, обусловленная интерференционными явлениями при ПРП
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_int_PRP[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_int_PRP[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_int_PRP[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_int_PRP[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sfi_tau_int_PRP[3]!="")
      {
      	//306	Коэффициент готовности в условиях интерференционных замираний при ПРП
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sfi_tau_int_PRP[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sfi_tau_int_PRP[1][0])+(unsigned char)(cDiD->sfi_tau_int_PRP[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sfi_tau_int_PRP[1][2])+(unsigned char)(cDiD->sfi_tau_int_PRP[1][3])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sfi_tau_int_PRP[1]).Right(7));
      	oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sfi_tau_int_PRP[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sfi_tau_int_PRP[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sSESR_PRP[3]!="")
      {
      	//307	%, Коэффициент секунд со значительным количеством ошибок при ПРП
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sSESR_PRP[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sSESR_PRP[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(4)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sSESR_PRP[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sSESR_PRP[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sK_ng_PRP[3]!="")
      {
      	//308	%, Коэффициент неготовности  при ПРП
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_PRP[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_PRP[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_PRP[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_PRP[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sSESR_norm[3]!="")
      {
      	//309	%, Норма на коэффициент секунд со значительным количеством ошибок при одинарном приеме
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sSESR_norm[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sSESR_norm[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(4)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sSESR_norm[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sSESR_norm[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sK_ng_norm[3]!="")
      {
      	//310	%, Норма на коэффициент неготовности при одинарном приеме
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_norm[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_norm[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_norm[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_norm[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }
   }
   //Краткий отчет
   else
   {


      //9	Тип оборудования
      //Объединяем ячейки
      oCell = oTable.Cell(N_strok+2,3);
   	oCell.Select();
   	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdExtend)));
   	oCells=oSel.GetCells();
   	oCells.Merge();
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sTyp_obor[0]);
      //Заполняем графы "Обозначение", "Размерн.", "Значение"
      oCell = oTable.Cell(N_strok+2,3);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sTyp_obor[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //10	Назначение РРЛ (Протяженность линии, км)
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sNaznachenie[0]);
      //Заполняем графы "Обозначение", "Размерн.", "Значение"
      oCell = oTable.Cell(N_strok+2,3);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sNaznachenie[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //12	Характер трассы(по влажности, по пересеченности, по высоте рельефа)
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sHarakter_trassy[0]);
      //Заполняем графы "Обозначение", "Размерн.", "Значение"
      oCell = oTable.Cell(N_strok+2,3);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sHarakter_trassy[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      if(cDiD->sPol[3]!="")
      {
      	//13	Поляризация
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oRan = oCell.GetRange();
      	oRan.SetText(cDiD->sPol[0]);
      	//Заполняем графы "Обозначение", "Размерн.", "Значение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oRan = oCell.GetRange();
      	oRan.SetText(cDiD->sPol[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      //14	Тип системы модуляции
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sTyp_sys[0]);
      //Заполняем графы "Обозначение", "Размерн.", "Значение"
      oCell = oTable.Cell(N_strok+2,3);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sTyp_sys[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //15	км, Протяженность интервала
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sR[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sR[1]);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sR[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sR[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //16	ГГц, Частотный диапазон
      //Добавляем строку
      oCell.Select();
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sf[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sf[1]);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sf[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sf[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //18	Мбит/с, Скорость передачи цифрового потока
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sC[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sC[1]);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sC[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sC[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //19	дБм, Мощность передатчика
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sP_pd[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sP_pd[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
   	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sP_pd[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sP_pd[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


       //23	дБм, Пороговая чувствительность приемника при BER
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sP_pm_por[0]);
      //Добавляем табулятор
      	oParFormat=oSel.GetParagraphFormat();
      	CTabStops oTabStops;
      	oTabStops=oParFormat.get_TabStops();
      	oTabStops.Add(80.2*MM2PH, COleVariant(short(wdAlignTabLeft)), COleVariant(short(wdTabLeaderSpaces)));
      	oParFormat.put_TabStops(oTabStops);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.SetText(cDiD->sBER[3]);//22	Коэффициент ошибок по битам, в зависимости от скорости передачи цифрового потока
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sP_pm_por[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
   	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sP_pm_por[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sP_pm_por[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //30	дБи, Коэффициент усиления основной антенны слева
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.SetText(cDiD->sG_pd[0]);
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.InsertSymbol(216, COleVariant("Arial"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      oSel.SetText(cDiD->sD_pd[3]);//29	м, Диаметр основной антенны слева
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sG_pd[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
   	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sG_pd[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sG_pd[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //32	дБи, Коэффициент усиления основной антенны справа
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.SetText(cDiD->sG_pm[0]);
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.InsertSymbol(216, COleVariant("Arial"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      oSel.SetText(cDiD->sD_pm[3]);//31	м, Диаметр основной антенны справа
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sG_pm[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
   	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sG_pm[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sG_pm[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      if(cDiD->sG1_dop[3]!="")
      {
      	//215	дБи, Коэффициент усиления дополнительной антенны слева (номинальное значение/с учетом ограничений)
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.SetText(cDiD->sG1_dop[0]);
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.InsertSymbol(216, COleVariant("Arial"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText(cDiD->sD1_dop[3]);//214	м, Диаметр дополнительной антенны слева
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sG1_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sG1_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sG1_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sG2_dop[3]!="")
      {
      	//217	дБи, Коэффициент усиления дополнительной антенны справа (номинальное значение/с учетом ограничений)
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.SetText(cDiD->sG2_dop[0]);
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.InsertSymbol(216, COleVariant("Arial"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText(cDiD->sD2_dop[3]);//216	м, Диаметр дополнительной антенны справа
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sG2_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sG2_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sG2_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      //39	м, Высота центра раскрыва основной антенны слева
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sH1[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sH1[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
   	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sH1[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sH1[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //40	м, Высота центра раскрыва основной антенны справа
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sH2[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sH2[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
   	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sH2[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sH2[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      if(cDiD->sH1_dop[3]!="")
      {
      	//212	м, Высота центра раскрыва дополнительной антенны слева
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sH2_dop[3]!="")
      {
      	//213	м, Высота центра раскрыва дополнительной антенны справа
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_dop[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_dop[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_dop[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      //43	1/м, Среднее значение градиента диэл. проницаемости воздуха
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sg_sr[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sg_sr[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
   	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sg_sr[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sg_sr[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //44	1/м, Стандартное отклонение градиента диэл. проницаемости воздуха
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sSigma[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.InsertSymbol((256*(unsigned char)(cDiD->sSigma[1][0])+(unsigned char)(cDiD->sSigma[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sSigma[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sSigma[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      if(cDiD->sSigma_ot_R[3]!="")
      {
      	//45	1/м, СКО градиента диэл. проницаемости воздуха в зависимости от расстояния
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sSigma_ot_R[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sSigma_ot_R[1][0])+(unsigned char)(cDiD->sSigma_ot_R[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sSigma_ot_R[1]).Right(3));
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sSigma_ot_R[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sSigma_ot_R[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      //46	дБ, Суммарные потери на интервале в антенно-фидерных трактах
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sEta_AFT[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.InsertSymbol((256*(unsigned char)(cDiD->sEta_AFT[1][0])+(unsigned char)(cDiD->sEta_AFT[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      oSel.SetText((cDiD->sEta_AFT[1]).Right(3));
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
   	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sEta_AFT[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sEta_AFT[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //47	дБ, Ослабление сигнала в свободном пространстве
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sW_0[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sW_0[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
   	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sW_0[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sW_0[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      if(cDiD->sV_difr_sr[3]!="")
      {
      	//48	дБ, Среднее ослабление за счет дифракции (при средней рефракции)
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_difr_sr[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_difr_sr[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_difr_sr[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_difr_sr[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sV_v_gazah[3]!="")
      {
      	//50	дБ, Среднее ослабление сигнала в газах атмосферы
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_v_gazah[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_v_gazah[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_v_gazah[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_v_gazah[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sV_50_proc[3]!="")
      {
      	//51	дБ, Среднее ослабление сигнала из-за перепада высот на горных и высокогорных трассах
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_50_proc[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_50_proc[1]);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_50_proc[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_50_proc[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      //52	дБм, Средний уровень сигнала на входе приемника
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sP_pm[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sP_pm[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
   	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sP_pm[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sP_pm[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      if(cDiD->sDelta_V_tip[3]!="")
      {
      	//53	дБ, Поправка на Vmin для типовых параметров аппаратуры
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_tip[0]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(13)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(3)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_V_tip[1][0])+(unsigned char)(cDiD->sDelta_V_tip[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sDelta_V_tip[1]).Right(5));
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_tip[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_tip[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sDelta_V_degr_int[3]!="")
      {
      	//81	дБ, Деградация порогового уровня из-за влияния помех на Tинт
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_degr_int[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_V_degr_int[1][0])+(unsigned char)(cDiD->sDelta_V_degr_int[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
         oSel.SetText((cDiD->sDelta_V_degr_int[1]).Right(10));
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_degr_int[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_degr_int[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sDelta_V_degr_subrefr[3]!="")
      {
      	//82	дБ, Деградация порогового уровня из-за влияния помех на T0
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_degr_subrefr[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_V_degr_subrefr[1][0])+(unsigned char)(cDiD->sDelta_V_degr_subrefr[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
         oSel.SetText((cDiD->sDelta_V_degr_subrefr[1]).Right(11));
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_degr_subrefr[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_degr_subrefr[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sDelta_V_degr_dojd[3]!="")
      {
      	//83	дБ, Деградация порогового уровня из-за влияния помех на Tд
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_degr_dojd[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_V_degr_dojd[1][0])+(unsigned char)(cDiD->sDelta_V_degr_dojd[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
         oSel.SetText((cDiD->sDelta_V_degr_dojd[1]).Right(11));
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_degr_dojd[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_degr_dojd[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sV_min_int[3]!="")
      {
      	//84	дБ, Минимально допустимый множитель интерференционного ослабления c учетом деградации порогового уровня, средних ослаблений и поправки на типовые параметры оборудования
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_int[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sV_min_int[1]);
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_int[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_int[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sV_min_subrefr[3]!="")
      {
      	//85	дБ, Минимально допустимый множитель субрефракционного ослабления c учетом деградации порогового уровня, средних ослаблений и поправки на типовые параметры оборудования
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_subrefr[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sV_min_subrefr[1]);
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_subrefr[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_subrefr[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sV_min_d[3]!="")
      {
      	//86	дБ, Минимально допустимый множитель ослабления в дожде c учетом деградации порогового уровня, средних ослаблений и поправки на типовые параметры оборудования
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_d[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sV_min_d[1]);
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_d[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_d[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sh_s[3]!="")
      {
      	//88	дБ, Высота сигнатуры
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sh_s[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sh_s[1]);
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sh_s[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sh_s[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sDelta_f_s[3]!="")
      {
      	//89	МГц, Ширина сигнатуры
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_f_s[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_f_s[1][0])+(unsigned char)(cDiD->sDelta_f_s[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
         oSel.SetText((cDiD->sDelta_f_s[1]).Right(2));
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_f_s[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_f_s[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sS_s[3]!="")
      {
      	//90	МГц, Площадь сигнатуры
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sS_s[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sS_s[1]);
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sS_s[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sS_s[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sI_ekv[3]!="")
      {
      	//91	дБ, Выигрыш за счет эквалайзера
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sI_ekv[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sI_ekv[1]);
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sI_ekv[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sI_ekv[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sV_min_eff_pred_ekv[3]!="")
      {
      	//93	дБ, Предельно реализуемый минимально допустимый эффективный множитель интерференционного ослабления с учетом эквалайзера
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_eff_pred_ekv[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sV_min_eff_pred_ekv[1]);
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_eff_pred_ekv[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_eff_pred_ekv[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sV_min_eff[3]!="")
      {
      	//94	дБ, Минимально допустимый эффективный множитель интерференционного ослабления
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_eff[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sV_min_eff[1]);
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_eff[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_eff[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      //97	км, Расстояние до критического препятствия при отсутствии рефракции
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sR_kr[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sR_kr[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sR_kr[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sR_kr[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //99	м, Просвет в точке критического препятствия при отсутствии рефракции
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sH_kr[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sH_kr[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sH_kr[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sH_kr[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //100	м, Оптимальный просвет в точке критического препятствия при отсутствии рефракции
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sH_0[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sH_0[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sH_0[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sH_0[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //101	км, Параметр хорды при отсутствии рефракции
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sParametr_r[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sParametr_r[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sParametr_r[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sParametr_r[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //105	Параметр, характеризующий кривизну аппроксимирующей сферы при отсутствии рефракции
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->smu_0[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.InsertSymbol((256*(unsigned char)(cDiD->smu_0[1][0])+(unsigned char)(cDiD->smu_0[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      oSel.SetText((cDiD->smu_0[1]).Right(1));
      oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->smu_0[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->smu_0[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //106	Параметр А для критического препятствия при отсутствии рефракции
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sA_kr[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sA_kr[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sA_kr[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sA_kr[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //114	Относительный просвет в точке критического препятствия при средней рефракции
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sp_ot_g_sr_kr[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sp_ot_g_sr_kr[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(3)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sp_ot_g_sr_kr[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sp_ot_g_sr_kr[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //119	1/м, Пороговое значение эффективного градиента диэлектрической проницаемости воздуха при Vдифр=Vмин.
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sg_0[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sg_0[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sg_0[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sg_0[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //145	Относительный просвет в точке критического препятствия при пороговой рефракции
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sp_ot_g_0_kr[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sp_ot_g_0_kr[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(3)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(3)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sp_ot_g_0_kr[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sp_ot_g_0_kr[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //149	Параметр, характеризующий кривизну аппроксимирующей сферы при пороговой рефракции
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->smu_3[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.InsertSymbol((256*(unsigned char)(cDiD->smu_3[1][0])+(unsigned char)(cDiD->smu_3[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      oSel.SetText((cDiD->smu_3[1]).Right(4));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->smu_3[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->smu_3[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //151	Параметр пси
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sPsi[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.InsertSymbol((256*(unsigned char)(cDiD->sPsi[1][0])+(unsigned char)(cDiD->sPsi[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sPsi[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sPsi[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //152	%, Неустойчивость, обусловленная рефракционными явлениями
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sT_0[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sT_0[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sT_0[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sT_0[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      if(cDiD->sR_otr_2[3]!="")
      {
      	//159	км, Расстояние до точки отражения для основных антенн при средней рефракции
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_otr_2[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_otr_2[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_otr_2[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_otr_2[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sH_otr_2[3]!="")
      {
      	//160	м, Просвет в точке отражения для основных антенн при средней рефракции
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_otr_2[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_otr_2[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_otr_2[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_otr_2[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sp_ot_g_sr_otr[3]!="")
      {
      	//162	Относительный просвет в точке отражения для основных антенн при средней рефракции
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sp_ot_g_sr_otr[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sp_ot_g_sr_otr[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(4)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(3)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sp_ot_g_sr_otr[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sp_ot_g_sr_otr[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sA_otr[3]!="")
      {
      	//165	Параметр А для точки отражения при средней рефракции
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sA_otr[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sA_otr[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sA_otr[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sA_otr[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sF_ot_p_g_A[3]!="")
      {
      	//166	Параметр F(p(g),A) для точки отражения при средней рефракции
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sF_ot_p_g_A[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sF_ot_p_g_A[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(5)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sF_ot_p_g_A[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sF_ot_p_g_A[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sK_vp[3]!="")
      {
      	//167	%, Коэффициент водной поверхности
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_vp[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_vp[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_vp[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_vp[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      //168	Параметр Q при средней рефракции
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sQ[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sQ[1]);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sQ[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sQ[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //169	%, Параметр, учитывающий вероятность возникновения многолучевых замираний, обусловленных отражениями радиоволн от слоистых неоднородностей тропосферы с перепадом диэлектрической проницаемости воздуха Delta_epsilon
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sT_Delta_eps[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText((cDiD->sT_Delta_eps[1]).Left(2));
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdMove)));
      oSel.InsertSymbol((256*(unsigned char)(cDiD->sT_Delta_eps[1][2])+(unsigned char)(cDiD->sT_Delta_eps[1][3])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      oSel.InsertSymbol((256*(unsigned char)(cDiD->sT_Delta_eps[1][4])+(unsigned char)(cDiD->sT_Delta_eps[1][5])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      oSel.SetText((cDiD->sT_Delta_eps[1]).Right(1));
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sT_Delta_eps[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sT_Delta_eps[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //173	%, Неустойчивость, обусловленная интерференционными явлениями
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sT_int[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sT_int[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sT_int[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sT_int[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      if(cDiD->sT_d_m[3]!="")
      {
      	//177	%, Неустойчивость, обусловленная влиянием осадков в наихудший месяц
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_d_m[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_d_m[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_d_m[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_d_m[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sQ_d[3]!="")
      {
      	//178	Коэффициент пересчета месячной статистики дождей к годовой
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sQ_d[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sQ_d[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sQ_d[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sQ_d[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sT_d_g[3]!="")
      {
      	//179	%, Неустойчивость, обусловленная влиянием осадков в среднем за год
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_d_g[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_d_g[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_d_g[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_d_g[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sfi_tau_0[3]!="")
      {
      	//188	Коэффициент готовности в условиях субрефракционных замираний
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sfi_tau_0[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sfi_tau_0[1][0])+(unsigned char)(cDiD->sfi_tau_0[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sfi_tau_0[1][2])+(unsigned char)(cDiD->sfi_tau_0[1][3])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sfi_tau_0[1]).Right(1));
      	oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sfi_tau_0[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sfi_tau_0[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sfi_tau_int[3]!="")
      {
      	//189	Коэффициент готовности в условиях интерференционных замираний
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sfi_tau_int[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sfi_tau_int[1][0])+(unsigned char)(cDiD->sfi_tau_int[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sfi_tau_int[1][2])+(unsigned char)(cDiD->sfi_tau_int[1][3])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sfi_tau_int[1]).Right(4));
      	oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sfi_tau_int[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sfi_tau_int[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      //190	Коэффициент интерференции
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sK_int[0]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(66)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(6)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sK_int[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sK_int[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sK_int[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //191	%, Коэффициент секунд со значительным количеством ошибок при одинарном приеме
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sSESR[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sSESR[1]);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sSESR[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sSESR[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      //192	%, Коэффициент неготовности при одинарном приеме
      //Добавляем строку
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //Заполняем графу "Наименование параметра"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sK_ng[0]);
      //Заполняем графу "Обозначение"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sK_ng[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //Заполняем графу "Размерн."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sK_ng[2]);
      //Заполняем графу "Значение"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sK_ng[3]);
      //Увеличиваем счетчик заполненных строк на 1
      N_strok++;


      if(cDiD->sK_stv[3]!="")
      {
      	//194	Количество рабочих стволов, не учитывая резервного
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_stv[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_stv[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_stv[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_stv[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sDelta_f[3]!="")
      {
      	//195	МГц, Разнос по частоте между резервным стволом и ближайшем к нему рабочим
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_f[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_f[1][0])+(unsigned char)(cDiD->sDelta_f[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sDelta_f[1]).Right(1));
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_f[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_f[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sSESR_ChRP[3]!="")
      {
      	//206	%, Коэффициент секунд со значительным количеством ошибок при ЧРП на интервале
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sSESR_ChRP[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sSESR_ChRP[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(4)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sSESR_ChRP[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sSESR_ChRP[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sK_ng_ChRP[3]!="")
      {
      	//207	%, Коэффициент неготовности на интервале ЦРРЛ с ЧРП в среднем за год
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_ChRP[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_ChRP[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_ChRP[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_ChRP[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sfi_tau_int_PRP[3]!="")
      {
      	//306	Коэффициент готовности в условиях интерференционных замираний при ПРП
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sfi_tau_int_PRP[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sfi_tau_int_PRP[1][0])+(unsigned char)(cDiD->sfi_tau_int_PRP[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sfi_tau_int_PRP[1][2])+(unsigned char)(cDiD->sfi_tau_int_PRP[1][3])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sfi_tau_int_PRP[1]).Right(7));
      	oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sfi_tau_int_PRP[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sfi_tau_int_PRP[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sSESR_PRP[3]!="")
      {
      	//307	%, Коэффициент секунд со значительным количеством ошибок при ПРП
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sSESR_PRP[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sSESR_PRP[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(4)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sSESR_PRP[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sSESR_PRP[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sK_ng_PRP[3]!="")
      {
      	//308	%, Коэффициент неготовности  при ПРП
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_PRP[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_PRP[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_PRP[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_PRP[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }


      if(cDiD->sSESR_norm[3]!="")
      {
      	//309	%, Норма на коэффициент секунд со значительным количеством ошибок при одинарном приеме
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sSESR_norm[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sSESR_norm[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(4)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sSESR_norm[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sSESR_norm[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }



      if(cDiD->sK_ng_norm[3]!="")
      {
      	//310	%, Норма на коэффициент неготовности при одинарном приеме
      	//Добавляем строку
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//Заполняем графу "Наименование параметра"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_norm[0]);
      	//Заполняем графу "Обозначение"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_norm[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//Заполняем графу "Размерн."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_norm[2]);
      	//Заполняем графу "Значение"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_norm[3]);
      	//Увеличиваем счетчик заполненных строк на 1
      	N_strok++;
      }
   }

   //Делаем нижнюю границу шапки отчета толстой линией
   oRow=oRows.Item(1);
   oBorders=oRow.GetBorders();
   oBorder=oBorders.Item(wdBorderBottom);
   oBorder.SetLineWidth(wdLineWidth150pt);

   return TRUE;
}//конец Draw_otchet_DMR(cData_interval_DMR &cDiD)

