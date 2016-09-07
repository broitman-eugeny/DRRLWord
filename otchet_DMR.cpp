//������� ������ ������ �� ����������� ������� ������������ ����������� otchet_DMR
#include <afx.h>
#include "afxdisp.h"//���������� ��� ������ AfxOleInit()
#include "msword9.h"//������������ ����, ���������� � ������� ClassWizard Visual Studio
#include "DMRWord.h"//������������ ���� � ���������� ������� MS Word



//������ �����������. ������������ ����� ������ ��� ������ ������� �� ������� ���� cData_interval_DMR
otchet_DMR::otchet_DMR()
{
}

//���������� ����� ����������� ������� ������������ ����������� ���� � ������� ��� ������ �����.
//��������� ����� �� ������ �� ������ ���� cData_interval_DMR, ��������� � "DMR.h"
//� ������ ������ ���������� TRUE, ����� - FALSE.
BOOL otchet_DMR::Draw_otchet_DMR(cData_interval_DMR *cDiD)
{
   COleVariant  covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);

   if(!Draw_blank(cDiD)) //���� �� ������� ���������� ����� � ������ � �������� ��������
   {
      //�������� Word
      app.Quit(COleVariant(short(wdDoNotSaveChanges)), COleVariant(short(NULL/*wdPromptUser*/)), COleVariant(short(false)));
   	return FALSE;
   }

   //��������� ���������� ��������
   CPageSetup oPageSetup;
   oPageSetup=Word_blank.GetPageSetup();
   oPageSetup.put_TopMargin(5.0*MM2PV);//���������� �� �������� ���� �����
   Word_blank.SetPageSetup(oPageSetup);

   //��������� ������ ��� �������
   _Font oFont;
   Selection oSel;
   oSel = app.GetSelection();
   oFont=oSel.GetFont();
   oFont.SetSize(1);
   oFont.SetName("Arial");


   //������� ����� ������� ������
   Tables oTables;
   Table oTable;
   Range oRan;
   oRan = oSel.GetRange();
   oTables = this->Word_blank.GetTables();
   //�������� ������� (1 ������, 5 ��������) � ���������
   oTable = oTables.Add(oRan,1,5,COleVariant(short(wdWord9TableBehavior)),COleVariant(short(wdAutoFitFixed)));
   //��������� ������� �� ��������
   Rows oRows;
   oRows=oTable.GetRows();
   oRows.SetLeftIndent(0.1);

   //��������� ������ � �������
   oTable.Select();
   oFont.SetSize(8);
   oFont.SetName("Arial");
   //������������� ����������� ���� ����� �������
	oTable.SetTopPadding(0.);
   oTable.SetBottomPadding(0.);
   oTable.SetLeftPadding(0.);
   oTable.SetRightPadding(0.);
   //������������� ����������� ���������� ����� �������� �������
   oTable.SetSpacing(0.);

   //������������� ������������ �������� ����� ��� ����������
   oTable.SetAllowAutoFit(BOOL(true));
   //������������� ������������ ������������ � ������� �������
   Cells oCells;
   oTable.Select();
   oCells=oSel.GetCells();
   oCells.SetVerticalAlignment(wdCellAlignVerticalCenter);
   //������������� ������ �����
   Row oRow;
   oRow=oRows.Item(1);
   oRow.SetHeight(4.2*MM2PV);

   //������������� ������ ��������
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

   //������������� ������� ������ ����� �������
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

   //����������� ������� � ������
   //����� "� �/�"
   Cell oCell;
   oCell = oTable.Cell(1,1);
   oRan = oCell.GetRange();
   oRan.SetText("� �/�");
   //����� "������������ ���������"
   oCell = oTable.Cell(1,2);
   oRan = oCell.GetRange();
   oRan.SetText("������������ ���������");

   //����� "�����������"
   oCell = oTable.Cell(1,3);
   oRan = oCell.GetRange();
   oRan.SetText("�����������");
   //����� "�������."
   oCell = oTable.Cell(1,4);
   oRan = oCell.GetRange();
   oRan.SetText("�������.");
   //����� "��������"
   oCell = oTable.Cell(1,5);
   oRan = oCell.GetRange();
   oRan.SetText("��������");
   //������������ ������ �� ������
   oTable.Select();
   Paragraphs oPars;
   oPars=oSel.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);


   //��������� ������
   oSel.SelectRow();
   oSel.InsertRowsBelow(COleVariant(short(1)));


   //��������� � ������ ������� ������������ ������
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


   //������������� ������������ �� ������ ���� � ����� "������������ ���������"
   oCell = oTable.Cell(2,2);
   oCell.Select();
   oPars=oSel.GetParagraphs();
   oPars.SetAlignment(AL_LEFT);

   //��������� ������
   oSel.SelectRow();
   oSel.InsertRowsBelow(COleVariant(short(1)));


   //��������� ������ � ������� "�����������"
   oCell = oTable.Cell(3,3);
   oCell.Select();
   oFont.SetSize(10);
   oFont.SetName("Times New Roman");


   //������������� ����� ������� � "���������"
   oRow=oRows.Item(1);
   oRow.SetHeadingFormat(True);

   int N_strok=0;//������� ����������� ����� �� ������� �����

   CParagraphFormat oParFormat;//��� �������������� ������������ � ������� �����������


   //������ �����
   if(cDiD->Vid_otcheta==POLNY)
   {


      //0	���� � ����� �������
      //���������� ������
      oCell = oTable.Cell(N_strok+2,3);
   	oCell.Select();
   	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdExtend)));
   	oCells=oSel.GetCells();
   	oCells.Merge();
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
   	oRan = oCell.GetRange();
   	oRan.SetText(cDiD->sCurDate[0]);

      //��������� ����� "�����������", "�������.", "��������"
      oCell = oTable.Cell(N_strok+2,3);
   	oRan = oCell.GetRange();
   	oRan.SetText(cDiD->sCurDate[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //1	���� � ��� ����� �������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sFileName[0]);
      //��������� ����� "�����������", "�������.", "��������"
      oCell = oTable.Cell(N_strok+2,3);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sFileName[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;

      //2	���� � ����� ���������� ��������� ����� �������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sFileDate[0]);
      //��������� ����� "�����������", "�������.", "��������"
      oCell = oTable.Cell(N_strok+2,3);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sFileDate[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;

      //3	�������� ������� ����� �� ����� �������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sLeftStationName[0]);
      //��������� ����� "�����������", "�������.", "��������"
      oCell = oTable.Cell(N_strok+2,3);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sLeftStationName[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;

      //4	�������� ������� ������ �� ����� �������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sRightStationName[0]);
      //��������� ����� "�����������", "�������.", "��������"
      oCell = oTable.Cell(N_strok+2,3);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sRightStationName[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;

      //5	����. ���., ������ ������ �� ����� �������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sForwardAzimuth[0]);
      //��������� ����� "�����������", "�������.", "��������"
      oCell = oTable.Cell(N_strok+2,3);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sForwardAzimuth[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;

      //6	����. ���., �������� ������ �� ����� �������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sBackwardAzimuth[0]);
      //��������� ����� "�����������", "�������.", "��������"
      oCell = oTable.Cell(N_strok+2,3);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sBackwardAzimuth[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;

      //7	�, ������� ������� ������� ����� �� ����� �������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sr_0[0]);
      //��������� ����� "�����������", "�������.", "��������"
      oCell = oTable.Cell(N_strok+2,3);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sr_0[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;

      //8	�, ������� ������� ������� ������ �� ����� �������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sr_R[0]);
      //��������� ����� "�����������", "�������.", "��������"
      oCell = oTable.Cell(N_strok+2,3);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sr_R[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //9	��� ������������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sTyp_obor[0]);

      //��������� ����� "�����������", "�������.", "��������"
      oCell = oTable.Cell(N_strok+2,3);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sTyp_obor[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //10	���������� ��� (������������� �����, ��)
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sNaznachenie[0]);
      //��������� ����� "�����������", "�������.", "��������"
      oCell = oTable.Cell(N_strok+2,3);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sNaznachenie[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //11	����� ���������������� ��� ����� �������������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sReconstr[0]);
      //��������� ����� "�����������", "�������.", "��������"
      oCell = oTable.Cell(N_strok+2,3);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sReconstr[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //12	�������� ������(�� ���������, �� ��������������, �� ������ �������)
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sHarakter_trassy[0]);
      //��������� ����� "�����������", "�������.", "��������"
      oCell = oTable.Cell(N_strok+2,3);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sHarakter_trassy[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      if(cDiD->sPol[3]!="")
      {
      	//13	�����������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oRan = oCell.GetRange();
      	oRan.SetText(cDiD->sPol[0]);
      	//��������� ����� "�����������", "�������.", "��������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oRan = oCell.GetRange();
      	oRan.SetText(cDiD->sPol[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      //14	��� ������� ���������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sTyp_sys[0]);
      //��������� ����� "�����������", "�������.", "��������"
      oCell = oTable.Cell(N_strok+2,3);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sTyp_sys[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //15	��, ������������� ���������
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sR[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sR[1]);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sR[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sR[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //16	���, ��������� ��������
      //��������� ������
      oCell.Select();
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sf[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sf[1]);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sf[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sf[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //17	�, ����� �����
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sLambda[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
   	oSel.InsertSymbol((256*(unsigned char)(cDiD->sLambda[1][0])+(unsigned char)(cDiD->sLambda[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sLambda[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sLambda[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //18	����/�, �������� �������� ��������� ������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sC[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sC[1]);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sC[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sC[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //19	���, �������� �����������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sP_pd[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sP_pd[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
   	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sP_pd[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sP_pd[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //21	���, ���������� ��������� ���������������� ��������� ��� �������� K_osh
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sP_pm_por_K_osh[0]);
      //��������� ���������
      	oParFormat=oSel.GetParagraphFormat();
      	CTabStops oTabStops;
      	oTabStops=oParFormat.get_TabStops();
      	oTabStops.Add(80.2*MM2PH, COleVariant(short(wdAlignTabLeft)), COleVariant(short(wdTabLeaderSpaces)));
      	oParFormat.put_TabStops(oTabStops);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.SetText(cDiD->sK_osh[3]);//20	����������� ������ ��� �������� ���������� ���������� ��������� ���������������� ���������

      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sP_pm_por_K_osh[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
   	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sP_pm_por_K_osh[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sP_pm_por_K_osh[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //23	���, ��������� ���������������� ��������� ��� BER
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sP_pm_por[0]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.SetText(cDiD->sBER[3]);//22	����������� ������ �� �����, � ����������� �� �������� �������� ��������� ������
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sP_pm_por[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
   	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sP_pm_por[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sP_pm_por[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //24	������� ��������� ������������ ������� �� ������ (����/���)
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sPass_retr[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sPass_retr[1]);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sPass_retr[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sPass_retr[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //25	��� ��� ����� (��������� �����/���������� ��� �������������� �����)
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sTruba_pd[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sTruba_pd[1]);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sTruba_pd[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sTruba_pd[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //26	��� ��� ������ (��������� �����/���������� ��� �������������� �����)
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sTruba_pm[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sTruba_pm[1]);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sTruba_pm[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sTruba_pm[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //27	�������� ������� ����� - ���������������? (��/���)
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sPeriskop_pd[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sPeriskop_pd[1]);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sPeriskop_pd[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sPeriskop_pd[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //28	�������� ������� ������ - ���������������? (��/���)
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sPeriskop_pm[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sPeriskop_pm[1]);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sPeriskop_pm[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sPeriskop_pm[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //30	���, ����������� �������� �������� ������� �����
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.SetText(cDiD->sG_pd[0]);
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.InsertSymbol(216, COleVariant("Arial"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      oSel.SetText(cDiD->sD_pd[3]);//29	�, ������� �������� ������� �����
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sG_pd[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
   	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sG_pd[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sG_pd[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //32	���, ����������� �������� �������� ������� ������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.SetText(cDiD->sG_pm[0]);
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.InsertSymbol(216, COleVariant("Arial"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      oSel.SetText(cDiD->sD_pm[3]);//31	�, ������� �������� ������� ������
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sG_pm[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
   	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sG_pm[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sG_pm[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      if(cDiD->sG1_dop[3]!="")
      {
      	//215	���, ����������� �������� �������������� ������� ����� (����������� ��������/� ������ �����������)
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.SetText(cDiD->sG1_dop[0]);
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.InsertSymbol(216, COleVariant("Arial"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText(cDiD->sD1_dop[3]);//214	�, ������� �������������� ������� �����
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sG1_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sG1_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sG1_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sG2_dop[3]!="")
      {
      	//217	���, ����������� �������� �������������� ������� ������ (����������� ��������/� ������ �����������)
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.SetText(cDiD->sG2_dop[0]);
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.InsertSymbol(216, COleVariant("Arial"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText(cDiD->sD2_dop[3]);//216	�, ������� �������������� ������� ������
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sG2_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sG2_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sG2_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      //33	��, ���������� ������ � ��� �����
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sEta_post_pd[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
   	oSel.InsertSymbol((256*(unsigned char)(cDiD->sEta_post_pd[1][0])+(unsigned char)(cDiD->sEta_post_pd[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      oSel.SetText((cDiD->sEta_post_pd[1]).Right(10));
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
   	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sEta_post_pd[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sEta_post_pd[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //34	��, ���������� ������ � ��� ������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sEta_post_pm[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
   	oSel.InsertSymbol((256*(unsigned char)(cDiD->sEta_post_pm[1][0])+(unsigned char)(cDiD->sEta_post_pm[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      oSel.SetText((cDiD->sEta_post_pm[1]).Right(11));
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
   	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sEta_post_pm[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sEta_post_pm[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      if(cDiD->sAlfa_AVT_pd[3]!="")
      {
      	//35	��/�, �������� ��������� ��������� �������� ������� �����
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sAlfa_AVT_pd[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.InsertSymbol((256*(unsigned char)(cDiD->sAlfa_AVT_pd[1][0])+(unsigned char)(cDiD->sAlfa_AVT_pd[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sAlfa_AVT_pd[1]).Right(9));
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sAlfa_AVT_pd[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sAlfa_AVT_pd[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sAlfa_AVT_pm[3]!="")
      {
      	//36	��/�, �������� ��������� ��������� �������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sAlfa_AVT_pm[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.InsertSymbol((256*(unsigned char)(cDiD->sAlfa_AVT_pm[1][0])+(unsigned char)(cDiD->sAlfa_AVT_pm[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sAlfa_AVT_pm[1]).Right(10));
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sAlfa_AVT_pm[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sAlfa_AVT_pm[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sL_AVT_pd[3]!="")
      {
      	//37	�, ����� ��������� �������� ������� �����
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sL_AVT_pd[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sL_AVT_pd[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sL_AVT_pd[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sL_AVT_pd[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sL_AVT_pm[3]!="")
      {
      	//38	�, ����� ��������� �������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sL_AVT_pm[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sL_AVT_pm[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sL_AVT_pm[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sL_AVT_pm[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      //39	�, ������ ������ �������� �������� ������� �����
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sH1[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sH1[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
   	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sH1[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sH1[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //40	�, ������ ������ �������� �������� ������� ������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sH2[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sH2[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
   	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sH2[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sH2[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      if(cDiD->sH1_dop[3]!="")
      {
      	//212	�, ������ ������ �������� �������������� ������� �����
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sH2_dop[3]!="")
      {
      	//213	�, ������ ������ �������� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      //41	�, ������� ������ ������ ���� ��� ������� ����
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sh_sr[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sh_sr[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
   	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sh_sr[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sh_sr[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //42	����� �������������� ������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sRaion[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sRaion[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
   	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sRaion[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sRaion[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //43	1/�, ������� �������� ��������� ����. ������������� �������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sg_sr[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sg_sr[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
   	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sg_sr[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sg_sr[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //44	1/�, ����������� ���������� ��������� ����. ������������� �������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sSigma[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.InsertSymbol((256*(unsigned char)(cDiD->sSigma[1][0])+(unsigned char)(cDiD->sSigma[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sSigma[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sSigma[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      if(cDiD->sSigma_ot_R[3]!="")
      {
      	//45	1/�, ��� ��������� ����. ������������� ������� � ����������� �� ����������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sSigma_ot_R[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sSigma_ot_R[1][0])+(unsigned char)(cDiD->sSigma_ot_R[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sSigma_ot_R[1]).Right(3));
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sSigma_ot_R[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sSigma_ot_R[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      //46	��, ��������� ������ �� ��������� � �������-�������� �������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sEta_AFT[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.InsertSymbol((256*(unsigned char)(cDiD->sEta_AFT[1][0])+(unsigned char)(cDiD->sEta_AFT[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      oSel.SetText((cDiD->sEta_AFT[1]).Right(3));
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
   	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sEta_AFT[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sEta_AFT[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //47	��, ���������� ������� � ��������� ������������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sW_0[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sW_0[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
   	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sW_0[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sW_0[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      if(cDiD->sV_difr_sr[3]!="")
      {
      	//48	��, ������� ���������� �� ���� ��������� (��� ������� ���������)
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_difr_sr[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_difr_sr[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_difr_sr[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_difr_sr[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sVlazh_para[3]!="")
      {
      	//49	�/�3, ��������� ����
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sVlazh_para[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sVlazh_para[1]);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sVlazh_para[2]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(3)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSuperscript(True);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sVlazh_para[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sV_v_gazah[3]!="")
      {
      	//50	��, ������� ���������� ������� � ����� ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_v_gazah[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_v_gazah[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_v_gazah[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_v_gazah[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sV_50_proc[3]!="")
      {
      	//51	��, ������� ���������� ������� ��-�� �������� ����� �� ������ � ������������ �������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_50_proc[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_50_proc[1]);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_50_proc[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_50_proc[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      //52	���, ������� ������� ������� �� ����� ���������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sP_pm[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sP_pm[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
   	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sP_pm[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sP_pm[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      if(cDiD->sDelta_V_tip[3]!="")
      {
      	//53	��, �������� �� Vmin ��� ������� ���������� ����������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_tip[0]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(13)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(3)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_V_tip[1][0])+(unsigned char)(cDiD->sDelta_V_tip[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sDelta_V_tip[1]).Right(5));
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_tip[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_tip[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sV_0_min_d[3]!="")
      {
      	//54	��, ���������� ���������� ��������� ���������� �� ������ ��� ����� ����������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_0_min_d[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_0_min_d[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_0_min_d[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_0_min_d[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sV_0_min_0[3]!="")
      {
      	//55	��, ���������� ���������� ��������� ���������� �� ������������ ��� ����� ����������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_0_min_0[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_0_min_0[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_0_min_0[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_0_min_0[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sZ_por_dop[3]!="")
      {
      	//56	��, ��������� ��������� (������/��� ���, ��� ���������� ���������� ������ �� 3 ��)
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_por_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_por_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_por_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_por_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sZ_por_dop_s[3]!="")
      {
      	//57	��, ��������� �������� ��������� ������� ��������� ������ � �������� ��������� �������, ���������� � ������ P��_���� ��� ���������� ���������� ������ �� 3 ��
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_por_dop_s[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_por_dop_s[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_por_dop_s[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_por_dop_s[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sZ_sosedn_kanal[3]!="")
      {
      	//58	��, ���������� ���������� (������/��� ���, ��� ���������� ���������� ������ �� 3 ��) ���������� ��������� (������ �� ��������� ������/��� ���)
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
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
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_sosedn_kanal[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_sosedn_kanal[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_sosedn_kanal[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sDelta_V_sosedn_kanal[3]!="")
      {
      	//59	��, ���������� ���������� ������ ��-�� ������� ����� �� ��������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_sosedn_kanal[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_V_sosedn_kanal[1][0])+(unsigned char)(cDiD->sDelta_V_sosedn_kanal[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sDelta_V_sosedn_kanal[1]).Right(2));
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_sosedn_kanal[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_sosedn_kanal[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      //60	������� "Co-channel" (����/���)
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sCochannel[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sCochannel[1]);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sCochannel[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sCochannel[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      if(cDiD->sD_p_0[3]!="")
      {
      	//61	��, ����������� ��������������� ������ ��� ���������� ��������� �������� �������� �������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sD_p_0[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sD_p_0[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sD_p_0[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sD_p_0[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sQ_p_cochannel[3]!="")
      {
      	//62	��, �����������, ����������� ������ �������������������� ��������� �������������� �������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sQ_p_cochannel[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sQ_p_cochannel[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sQ_p_cochannel[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sQ_p_cochannel[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sD_p[3]!="")
      {
      	//63	��, ����������� ��������������� ������ � �������� ����������������� ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sD_p[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sD_p[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sD_p[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sD_p[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sD_p_dojd[3]!="")
      {
      	//64	��, ����������� ��������������� ������, ������������� �������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sD_p_dojd[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sD_p_dojd[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sD_p_dojd[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sD_p_dojd[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sI_xpic[3]!="")
      {
      	//65	��, ������� ������������ �������������������� �����
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sI_xpic[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sI_xpic[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sI_xpic[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sI_xpic[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sZ_polyariz_cochannel[3]!="")
      {
      	//66	��, ���������� ���������� (�������������������� ������/��� ���) ���������� ��������� (������/��� ���, ��� ���������� ���������� ������ �� 3 ��), ��� ����� � �������� ���� (� �������� �0 �������������������� ������ �� �����������)
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
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
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_polyariz_cochannel[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_polyariz_cochannel[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_polyariz_cochannel[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sZ_polyariz_cochannel_dojd[3]!="")
      {
      	//67	��, ���������� ���������� (�������������������� ������/��� ���) ���������� ��������� (������/��� ���, ��� ���������� ���������� ������ �� 3 ��), ��� ����� � �������� �� (� �������� �0 �������������������� ������ �� �����������)
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
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
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_polyariz_cochannel_dojd[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_polyariz_cochannel_dojd[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_polyariz_cochannel_dojd[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sDelta_V_polyariz_cochannel[3]!="")
      {
      	//68	��, ���������� ���������� ������ ��-�� ������� �������������������� ����� ��� ����� � �������� ���� (� �������� �0 �������������������� ������ �� �����������)
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_polyariz_cochannel[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_V_polyariz_cochannel[1][0])+(unsigned char)(cDiD->sDelta_V_polyariz_cochannel[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sDelta_V_polyariz_cochannel[1]).Right(7));
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_polyariz_cochannel[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_polyariz_cochannel[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sDelta_V_polyariz_cochannel_dojd[3]!="")
      {
      	//69	��, ���������� ���������� ������ ��-�� ������� �������������������� ����� ��� ����� � �������� �� (� �������� �0 �������������������� ������ �� �����������)
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_polyariz_cochannel_dojd[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_V_polyariz_cochannel_dojd[1][0])+(unsigned char)(cDiD->sDelta_V_polyariz_cochannel_dojd[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sDelta_V_polyariz_cochannel_dojd[1]).Right(5));
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_polyariz_cochannel_dojd[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_polyariz_cochannel_dojd[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sN_mesh[3]!="")
      {
      	//70	���������� �������� ����������, ���������� �� ������� ��������� �������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sN_mesh[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sN_mesh[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sN_mesh[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sN_mesh[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sR_mesh[3]!="")
      {
      	//71	��, ������������� �������� ���������� (����� ����)
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_mesh[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_mesh[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_mesh[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_mesh[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sF_ot_alfa[3]!="")
      {
      	//72	��, ���������� �������� �������� �� ���� �������� �������������� ������ (����� ����)
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sF_ot_alfa[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.SetText((cDiD->sF_ot_alfa[1]).Left(2));
      	oSel.Collapse(COleVariant(short(wdCollapseEnd)));
         oSel.InsertSymbol((256*(unsigned char)(cDiD->sF_ot_alfa[1][2])+(unsigned char)(cDiD->sF_ot_alfa[1][3])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
         oSel.SetText((cDiD->sF_ot_alfa[1]).Right(1));
         //��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sF_ot_alfa[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sF_ot_alfa[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sP_pd_mesh[3]!="")
      {
      	//73	���, �������� ������������ �������� ���������� (����� ����)
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sP_pd_mesh[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sP_pd_mesh[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sP_pd_mesh[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sP_pd_mesh[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sARM[3]!="")
      {
      	//74	��, �������� �������������� ����������� �������� ������������ �������� ���������� (����� ����)
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sARM[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sARM[1]);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sARM[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sARM[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sV_sr_mesh[3]!="")
      {
      	//75	��, ������� ���������� �� ������� �������� ���������� (����� ����)
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_sr_mesh[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sV_sr_mesh[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_sr_mesh[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_sr_mesh[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sG_pd_mesh[3]!="")
      {
      	//76	���, ������������ �������� ���������� ������ �������� ���������� (����� ����)
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sG_pd_mesh[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sG_pd_mesh[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sG_pd_mesh[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sG_pd_mesh[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sL_AVT_pd_mesh[3]!="")
      {
      	//77	�, ����� ���������� ������������ �������� ���������� (����� ����)
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sL_AVT_pd_mesh[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sL_AVT_pd_mesh[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sL_AVT_pd_mesh[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sL_AVT_pd_mesh[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sZ_por[3]!="")
      {
      	//78	��, ��������� (��������� ������ ��������� ����������� � ���������������/��� ���)
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_por[0]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(67)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sZ_por[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_por[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_por[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sZ_obr_uzl[3]!="")
      {
      	//79	��, ���������� ���������� (��������� ������ ��������� ����������� � ���������������/��� ���) ���������� ��������� (������/��� ���, ��� ���������� ���������� ������ �� 3 ��)
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
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
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sZ_obr_uzl[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_obr_uzl[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_obr_uzl[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sDelta_V_obr_uzl[3]!="")
      {
      	//80	��, ���������� ���������� ������ ��-�� ������� ����� � ��������� ����������� � ���������������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_obr_uzl[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_V_obr_uzl[1][0])+(unsigned char)(cDiD->sDelta_V_obr_uzl[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
         oSel.SetText((cDiD->sDelta_V_obr_uzl[1]).Right(9));
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_obr_uzl[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_obr_uzl[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sDelta_V_degr_int[3]!="")
      {
      	//81	��, ���������� ���������� ������ ��-�� ������� ����� �� T���
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_degr_int[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_V_degr_int[1][0])+(unsigned char)(cDiD->sDelta_V_degr_int[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
         oSel.SetText((cDiD->sDelta_V_degr_int[1]).Right(10));
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_degr_int[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_degr_int[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sDelta_V_degr_subrefr[3]!="")
      {
      	//82	��, ���������� ���������� ������ ��-�� ������� ����� �� T0
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_degr_subrefr[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_V_degr_subrefr[1][0])+(unsigned char)(cDiD->sDelta_V_degr_subrefr[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
         oSel.SetText((cDiD->sDelta_V_degr_subrefr[1]).Right(11));
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_degr_subrefr[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_degr_subrefr[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sDelta_V_degr_dojd[3]!="")
      {
      	//83	��, ���������� ���������� ������ ��-�� ������� ����� �� T�
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_degr_dojd[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_V_degr_dojd[1][0])+(unsigned char)(cDiD->sDelta_V_degr_dojd[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
         oSel.SetText((cDiD->sDelta_V_degr_dojd[1]).Right(11));
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_degr_dojd[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_degr_dojd[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sV_min_int[3]!="")
      {
      	//84	��, ���������� ���������� ��������� ������������������ ���������� c ������ ���������� ���������� ������, ������� ���������� � �������� �� ������� ��������� ������������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_int[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sV_min_int[1]);
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_int[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_int[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sV_min_subrefr[3]!="")
      {
      	//85	��, ���������� ���������� ��������� ����������������� ���������� c ������ ���������� ���������� ������, ������� ���������� � �������� �� ������� ��������� ������������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_subrefr[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sV_min_subrefr[1]);
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_subrefr[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_subrefr[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sV_min_d[3]!="")
      {
      	//86	��, ���������� ���������� ��������� ���������� � ����� c ������ ���������� ���������� ������, ������� ���������� � �������� �� ������� ��������� ������������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_d[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sV_min_d[1]);
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_d[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_d[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sEkv[3]!="")
      {
      	//87	������� � ������� ����������� (����/���)
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sEkv[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sEkv[1]);
         //��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sEkv[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sEkv[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sh_s[3]!="")
      {
      	//88	��, ������ ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sh_s[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sh_s[1]);
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sh_s[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sh_s[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sDelta_f_s[3]!="")
      {
      	//89	���, ������ ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_f_s[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_f_s[1][0])+(unsigned char)(cDiD->sDelta_f_s[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
         oSel.SetText((cDiD->sDelta_f_s[1]).Right(2));
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_f_s[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_f_s[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sS_s[3]!="")
      {
      	//90	���, ������� ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sS_s[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sS_s[1]);
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sS_s[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sS_s[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sI_ekv[3]!="")
      {
      	//91	��, ������� �� ���� �����������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sI_ekv[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sI_ekv[1]);
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sI_ekv[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sI_ekv[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sV_min_eff_pred[3]!="")
      {
      	//92	��, ��������� ����������� ���������� ���������� ����������� ��������� ������������������ ���������� ��� ����� �����������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_eff_pred[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sV_min_eff_pred[1]);
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_eff_pred[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_eff_pred[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sV_min_eff_pred_ekv[3]!="")
      {
      	//93	��, ��������� ����������� ���������� ���������� ����������� ��������� ������������������ ���������� � ������ �����������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_eff_pred_ekv[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sV_min_eff_pred_ekv[1]);
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_eff_pred_ekv[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_eff_pred_ekv[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sV_min_eff[3]!="")
      {
      	//94	��, ���������� ���������� ����������� ��������� ������������������ ����������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_eff[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sV_min_eff[1]);
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_eff[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_eff[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sR_pr[3]!="")
      {
      	//95	��, ���������� �� ������������ ����������� �� ������� �������� �������������� �������� ��� ���������� ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_pr[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_pr[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_pr[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_pr[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sH_pr[3]!="")
      {
      	//96	�, ������� � ����� ������������ ����������� �� ������� �������� �������������� �������� ��� ���������� ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_pr[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_pr[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_pr[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_pr[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      //97	��, ���������� �� ������������ ����������� ��� ���������� ���������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sR_kr[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sR_kr[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sR_kr[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sR_kr[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //98	������������� ���������� ������������ ����������� ��� ���������� ���������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sk[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sk[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sk[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sk[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //99	�, ������� � ����� ������������ ����������� ��� ���������� ���������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sH_kr[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sH_kr[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sH_kr[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sH_kr[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //100	�, ����������� ������� � ����� ������������ ����������� ��� ���������� ���������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sH_0[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sH_0[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sH_0[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sH_0[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //101	��, �������� ����� ��� ���������� ���������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sParametr_r[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sParametr_r[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sParametr_r[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sParametr_r[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //102	��������� ����� ����� � ������������� ��������� ��� ���������� ���������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->slr[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->slr[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->slr[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->slr[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //103	�, ������ �������� ���������������� ����� ��� ���������� ���������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sDelta_y[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_y[1][0])+(unsigned char)(cDiD->sDelta_y[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      oSel.SetText((cDiD->sDelta_y[1]).Right(5));
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sDelta_y[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sDelta_y[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //104	��������� ������ ���������������� ����� � ������������ �������� ��� ���������� ���������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sAlfa_delta_y[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.InsertSymbol((256*(unsigned char)(cDiD->sAlfa_delta_y[1][0])+(unsigned char)(cDiD->sAlfa_delta_y[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      oSel.InsertSymbol((256*(unsigned char)(cDiD->sAlfa_delta_y[1][2])+(unsigned char)(cDiD->sAlfa_delta_y[1][3])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      oSel.SetText((cDiD->sAlfa_delta_y[1]).Right(5));
      oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sAlfa_delta_y[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sAlfa_delta_y[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //105	��������, ��������������� �������� ���������������� ����� ��� ���������� ���������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->smu_0[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.InsertSymbol((256*(unsigned char)(cDiD->smu_0[1][0])+(unsigned char)(cDiD->smu_0[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      oSel.SetText((cDiD->smu_0[1]).Right(1));
      oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->smu_0[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->smu_0[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //106	�������� � ��� ������������ ����������� ��� ���������� ���������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sA_kr[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sA_kr[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sA_kr[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sA_kr[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //107	��, ������������� ������ ����� ��� ������� ���������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sa_e_km_2[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sa_e_km_2[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sa_e_km_2[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sa_e_km_2[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      if(cDiD->sR_pr_2[3]!="")
      {
      	//108	��, ���������� �� ������������ ����������� �� ������� �������� �������������� �������� ��� ������� ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_pr_2[0]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(64)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_pr_2[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_pr_2[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_pr_2[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sH_pr_2[3]!="")
      {
      	//109	�, ������� � ����� ������������ ����������� �� ������� �������� �������������� �������� ��� ������� ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_pr_2[0]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(63)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_pr_2[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_pr_2[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_pr_2[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      //110	��, ���������� �� ������������ ����������� ��� ������� ���������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sR_kr_2[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sR_kr_2[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sR_kr_2[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sR_kr_2[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //111	������������� ���������� ������������ ����������� ��� ������� ���������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sk_2[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sk_2[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sk_2[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sk_2[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //112	�, ������� � ����� ������������ ����������� ��� ������� ���������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sH_kr_2[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sH_kr_2[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sH_kr_2[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sH_kr_2[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //113	�, ����������� ������� � ����� ������������ ����������� ��� ������� ���������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sH_0_2[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sH_0_2[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sH_0_2[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sH_0_2[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //114	������������� ������� � ����� ������������ ����������� ��� ������� ���������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sp_ot_g_sr_kr[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sp_ot_g_sr_kr[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(3)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sp_ot_g_sr_kr[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sp_ot_g_sr_kr[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //115	��, �������� ����� ��� ������� ���������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sParametr_r_2[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sParametr_r_2[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sParametr_r_2[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sParametr_r_2[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //116	�, ������ �������� ���������������� ����� ��� ������� ���������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sDelta_y_2[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_y_2[1][0])+(unsigned char)(cDiD->sDelta_y_2[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      oSel.SetText((cDiD->sDelta_y_2[1]).Right(6));
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sDelta_y_2[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sDelta_y_2[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //117	k�, ������ ���������������� ����� ��� ������� ���������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sb_2[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sb_2[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sb_2[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sb_2[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //118	��������, ��������������� �������� ���������������� ����� ��� ������� ���������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->smu_2[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.InsertSymbol((256*(unsigned char)(cDiD->smu_2[1][0])+(unsigned char)(cDiD->smu_2[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      oSel.SetText((cDiD->smu_2[1]).Right(5));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->smu_2[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->smu_2[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //119	1/�, ��������� �������� ������������ ��������� ��������������� ������������� ������� ��� V����=V���.
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sg_0[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sg_0[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sg_0[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sg_0[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //120	��, ������������� ������ ����� ��� ��������� ���������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sa_e_km_3[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sa_e_km_3[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sa_e_km_3[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sa_e_km_3[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //121	��, ���������� �� ����� ������� ����� ��� ��������� ���������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sR1x_3[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sR1x_3[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sR1x_3[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sR1x_3[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //122	��, ���������� �� ������ ������� ����� ��� ��������� ���������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sR2x_3[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sR2x_3[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sR2x_3[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sR2x_3[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //123	�, �������� ������� ����� ������� ����� ��� ��������� ���������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sH1h_3[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sH1h_3[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sH1h_3[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sH1h_3[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //124	�, �������� ������� ������ ������� ����� ��� ��������� ���������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sH2h_3[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sH2h_3[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sH2h_3[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sH2h_3[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //125	�, �������� ������� ����� � ����� �������� �������� ����������� �� �������� ������������ �������������� �������� ��� ��������� ���������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sH_kr_h_3[0]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(59)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sH_kr_h_3[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sH_kr_h_3[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sH_kr_h_3[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      if(cDiD->sR1_kas_3[3]!="")
      {
      	//126	��, ���������� �� ����� ������� ����� ����������� � ������������ ������� ��� ��������� ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sR1_kas_3[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sR1_kas_3[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sR1_kas_3[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sR1_kas_3[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sR2_kas_3[3]!="")
      {
      	//127	��, ���������� �� ����� ������� ������ ����������� � ������������ ������� ��� ��������� ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sR2_kas_3[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sR2_kas_3[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sR2_kas_3[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sR2_kas_3[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sH1_kas_3[3]!="")
      {
      	//128	�, �������� ������� ����� ������� ����� ����������� � ������������ ������� ��� ��������� ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_kas_3[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_kas_3[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_kas_3[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_kas_3[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sH2_kas_3[3]!="")
      {
      	//129	�, �������� ������� ����� ������� ������ ����������� � ������������ ������� ��� ��������� ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_kas_3[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_kas_3[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_kas_3[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_kas_3[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sR1_ellipse_3[3]!="")
      {
      	//130	��, ���������� �� ����� ����������� ������ ����������� � ������������ ������� ��� ��������� ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sR1_ellipse_3[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sR1_ellipse_3[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sR1_ellipse_3[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sR1_ellipse_3[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sR2_ellipse_3[3]!="")
      {
      	//131	��, ���������� �� ����� ����������� ������� ����������� � ������������ ������� ��� ��������� ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sR2_ellipse_3[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sR2_ellipse_3[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sR2_ellipse_3[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sR2_ellipse_3[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sH1_ellipse_3[3]!="")
      {
      	//132	�, �������� ������� ����� ����������� ������ ����������� � ������������ ������� ��� ��������� ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_ellipse_3[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_ellipse_3[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_ellipse_3[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_ellipse_3[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sH2_ellipse_3[3]!="")
      {
      	//133	�, �������� ������� ����� ����������� ������� ����������� � ������������ ������� ��� ��������� ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_ellipse_3[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_ellipse_3[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_ellipse_3[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_ellipse_3[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sR1_horda_H0_3[3]!="")
      {
      	//134	��, ���������� �� ����� ������� �����, ����������� �� �������� H0, ��� ��������� ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
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
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sR1_horda_H0_3[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sR1_horda_H0_3[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sR1_horda_H0_3[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sR2_horda_H0_3[3]!="")
      {
      	//135	��, ���������� �� ������ ������� �����, ����������� �� �������� H0, ��� ��������� ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
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
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sR2_horda_H0_3[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sR2_horda_H0_3[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sR2_horda_H0_3[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sH1_horda_H0_3[3]!="")
      {
      	//136	�, �������� ������� ����� ������� �����, ����������� �� �������� H0, ��� ��������� ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
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
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_horda_H0_3[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_horda_H0_3[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_horda_H0_3[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sH2_horda_H0_3[3]!="")
      {
      	//137	�, �������� ������� ������ ������� �����, ����������� �� �������� H0, ��� ��������� ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
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
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_horda_H0_3[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_horda_H0_3[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_horda_H0_3[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sR_cross_kas_3[3]!="")
      {
      	//138	��, ���������� �� ����� ����������� ����������� � ����������� ������� ��� ��������� ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_cross_kas_3[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_cross_kas_3[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_cross_kas_3[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_cross_kas_3[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sH_cross_kas_3[3]!="")
      {
      	//139	�, �������� ������� ����� ����������� ����������� � ����������� ������� ��� ��������� ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_cross_kas_3[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_cross_kas_3[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_cross_kas_3[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_cross_kas_3[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sProsvet_cross_kas_3[3]!="")
      {
      	//140	�, ������� � ����� ����������� ����������� ��� ��������� ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sProsvet_cross_kas_3[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sProsvet_cross_kas_3[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sProsvet_cross_kas_3[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sProsvet_cross_kas_3[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      //141	��, ���������� �� ������������ ����������� ��� ��������� ���������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sR_kr_3[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sR_kr_3[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sR_kr_3[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sR_kr_3[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //142	������������� ���������� ������������ ����������� ��� ��������� ���������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sk_3[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sk_3[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sk_3[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sk_3[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //143	�, ������� � ����� ������������ ����������� ��� ��������� ���������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sH_kr_3[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sH_kr_3[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sH_kr_3[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sH_kr_3[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //144	�, ����������� ������� � ����� ������������ ����������� ��� ��������� ���������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sH_0_3[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sH_0_3[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sH_0_3[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sH_0_3[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //145	������������� ������� � ����� ������������ ����������� ��� ��������� ���������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sp_ot_g_0_kr[0]);
      //��������� ����� "�����������"
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
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sp_ot_g_0_kr[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sp_ot_g_0_kr[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //146	��, �������� ����� ��� ��������� ���������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sParametr_r_3[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sParametr_r_3[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sParametr_r_3[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sParametr_r_3[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //147	�, ������ �������� ���������������� ����� ��� ��������� ���������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sDelta_y_3[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_y_3[1][0])+(unsigned char)(cDiD->sDelta_y_3[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      oSel.SetText((cDiD->sDelta_y_3[1]).Right(5));
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sDelta_y_3[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sDelta_y_3[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //148	k�, ������ ���������������� ����� ��� ��������� ���������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sb_3[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sb_3[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sb_3[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sb_3[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //149	��������, ��������������� �������� ���������������� ����� ��� ��������� ���������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->smu_3[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.InsertSymbol((256*(unsigned char)(cDiD->smu_3[1][0])+(unsigned char)(cDiD->smu_3[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      oSel.SetText((cDiD->smu_3[1]).Right(4));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->smu_3[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->smu_3[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //150	��, ������������� ���������� ��� ��������� ���������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sV_difr_ot_g_0[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sV_difr_ot_g_0[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sV_difr_ot_g_0[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sV_difr_ot_g_0[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //151	�������� ���
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sPsi[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.InsertSymbol((256*(unsigned char)(cDiD->sPsi[1][0])+(unsigned char)(cDiD->sPsi[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sPsi[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sPsi[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //152	%, ��������������, ������������� �������������� ���������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sT_0[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sT_0[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sT_0[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sT_0[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      if(cDiD->sR_otr[3]!="")
      {
      	//153	��, ���������� ��������� ����� ��������� ��� ���������� ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_otr[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_otr[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_otr[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_otr[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sH_otr[3]!="")
      {
      	//154	�, ������� � ��������� ����� ��������� ��� ���������� ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_otr[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_otr[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_otr[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_otr[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sH_0_otr[3]!="")
      {
      	//155	�, ����������� ������� � ��������� ����� ��������� ��� ���������� ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_0_otr[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_0_otr[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_0_otr[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_0_otr[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sp_ot_g_otr[3]!="")
      {
      	//156	������������� ������� � ��������� ����� ��������� ��� ���������� ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sp_ot_g_otr[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sp_ot_g_otr[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(4)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sp_ot_g_otr[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sp_ot_g_otr[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sX_prodoln[3]!="")
      {
      	//157	��, ������������� ����������� ������� �������, ��������� ��� ������������������ ��������� ������ ��� ���������� ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sX_prodoln[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sX_prodoln[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sX_prodoln[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sX_prodoln[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sDelta_h_max[3]!="")
      {
      	//158	�, ������������ ���������� ������� ����������� ������� ������� �� ���������������� �����������, ��� ������� ��� ����� �������, ��� ������� �����������������, ��� ���������� ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_h_max[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_h_max[1][0])+(unsigned char)(cDiD->sDelta_h_max[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
         oSel.SetText((cDiD->sDelta_h_max[1]).Right(10));
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_h_max[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_h_max[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sR_otr_2[3]!="")
      {
      	//159	��, ���������� �� ����� ��������� ��� �������� ������ ��� ������� ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_otr_2[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_otr_2[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_otr_2[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_otr_2[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sH_otr_2[3]!="")
      {
      	//160	�, ������� � ����� ��������� ��� �������� ������ ��� ������� ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_otr_2[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_otr_2[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_otr_2[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_otr_2[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sH_0_otr_2[3]!="")
      {
      	//161	�, ����������� ������� � ����� ��������� ��� �������� ������ ��� ������� ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_0_otr_2[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_0_otr_2[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_0_otr_2[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_0_otr_2[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sp_ot_g_sr_otr[3]!="")
      {
      	//162	������������� ������� � ����� ��������� ��� �������� ������ ��� ������� ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sp_ot_g_sr_otr[0]);
      	//��������� ����� "�����������"
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
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sp_ot_g_sr_otr[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sp_ot_g_sr_otr[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sX_prodoln_2[3]!="")
      {
      	//163	��, ������������� ����������� ������� �������, ��������� ��� ������������������ ��������� ������, ��� ������� ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sX_prodoln_2[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sX_prodoln_2[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sX_prodoln_2[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sX_prodoln_2[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sDelta_h_max_2[3]!="")
      {
      	//164	�, ������������ ���������� ������� ����������� ������� ������� �� ���������������� �����������, ��� ������� ��� ����� �������, ��� ������� �����������������, ��� ������� ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_h_max_2[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_h_max_2[1][0])+(unsigned char)(cDiD->sDelta_h_max_2[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
         oSel.SetText((cDiD->sDelta_h_max_2[1]).Right(11));
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_h_max_2[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_h_max_2[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sA_otr[3]!="")
      {
      	//165	�������� � ��� ����� ��������� ��� ������� ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sA_otr[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sA_otr[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sA_otr[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sA_otr[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sF_ot_p_g_A[3]!="")
      {
      	//166	�������� F(p(g),A) ��� ����� ��������� ��� ������� ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sF_ot_p_g_A[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sF_ot_p_g_A[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(5)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sF_ot_p_g_A[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sF_ot_p_g_A[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sK_vp[3]!="")
      {
      	//167	%, ����������� ������ �����������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_vp[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_vp[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_vp[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_vp[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      //168	�������� Q ��� ������� ���������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sQ[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sQ[1]);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sQ[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sQ[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //169	%, ��������, ����������� ����������� ������������� ������������ ���������, ������������� ����������� ��������� �� �������� ��������������� ���������� � ��������� ��������������� ������������� ������� Delta_epsilon
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sT_Delta_eps[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText((cDiD->sT_Delta_eps[1]).Left(2));
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdMove)));
      oSel.InsertSymbol((256*(unsigned char)(cDiD->sT_Delta_eps[1][2])+(unsigned char)(cDiD->sT_Delta_eps[1][3])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      oSel.InsertSymbol((256*(unsigned char)(cDiD->sT_Delta_eps[1][4])+(unsigned char)(cDiD->sT_Delta_eps[1][5])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      oSel.SetText((cDiD->sT_Delta_eps[1]).Right(1));
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sT_Delta_eps[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sT_Delta_eps[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      if(cDiD->ssigma_0[3]!="")
      {
      	//170	��, ����������� ���������� �������������� ����������� ������ ������������� V � ������� T(V)>1%
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->ssigma_0[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->ssigma_0[1][0])+(unsigned char)(cDiD->ssigma_0[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->ssigma_0[1]).Right(1));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->ssigma_0[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->ssigma_0[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->ssigma_1[3]!="")
      {
      	//171	��, ����������� ���������� �������������� ����������� ������ ������������� V � ������� T(V)<1%
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->ssigma_1[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->ssigma_1[1][0])+(unsigned char)(cDiD->ssigma_1[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->ssigma_1[1]).Right(1));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->ssigma_1[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->ssigma_1[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sT_int_priz[3]!="")
      {
      	//172.1	%, ��������������, ������������� ������������������ ��������� �� ��������� �������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_int_priz[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_int_priz[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_int_priz[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_int_priz[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }

      if(cDiD->sT_int_vg[3]!="")
      {
      	//172.2	%, ��������������, ������������� ������������������ ��������� �� ������������ �������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_int_vg[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_int_vg[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_int_vg[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_int_vg[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      //173	%, ��������������, ������������� ������������������ ���������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sT_int[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sT_int[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sT_int[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sT_int[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      if(cDiD->sRaion_d[3]!="")
      {
      	//174	����� ������ �� ������������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sRaion_d[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sRaion_d[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sRaion_d[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sRaion_d[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sRaion_Qd[3]!="")
      {
      	//175	����� ������ ������������� Q�
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sRaion_Qd[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sRaion_Qd[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sRaion_Qd[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sRaion_Qd[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sI_max[3]!="")
      {
      	//176	��/�, ����������� ���������� ������������� �����
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sI_max[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sI_max[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sI_max[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sI_max[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sT_d_m[3]!="")
      {
      	//177	%, ��������������, ������������� �������� ������� � ��������� �����
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_d_m[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_d_m[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_d_m[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_d_m[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sQ_d[3]!="")
      {
      	//178	����������� ��������� �������� ���������� ������ � �������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sQ_d[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sQ_d[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sQ_d[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sQ_d[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sT_d_g[3]!="")
      {
      	//179	%, ��������������, ������������� �������� ������� � ������� �� ���
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_d_g[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_d_g[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_d_g[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_d_g[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sPsi_tau_0[3]!="")
      {
      	//180	��^2, ���������� �������� ��� ����������� C_m_0
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sPsi_tau_0[0]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(37)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sPsi_tau_0[1][0])+(unsigned char)(cDiD->sPsi_tau_0[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sPsi_tau_0[1][2])+(unsigned char)(cDiD->sPsi_tau_0[1][3])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sPsi_tau_0[1]).Right(5));
      	oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sPsi_tau_0[2]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSuperscript(True);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sPsi_tau_0[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sC_m_0[3]!="")
      {
      	//181	�, ������������ ����������� ��� ������� tau_m_0
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_m_0[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_m_0[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_m_0[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_m_0[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->stau_m_0[3]!="")
      {
      	//182	�, ��������� �������� ������������ ��������� � �������� ���������������� ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->stau_m_0[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->stau_m_0[1][0])+(unsigned char)(cDiD->stau_m_0[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->stau_m_0[1]).Right(7));
      	oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->stau_m_0[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->stau_m_0[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->ssigma_tau_0[3]!="")
      {
      	//183	��, ����������� ���������� ��� ��������� ������������ ��������� � �������� �������. ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->ssigma_tau_0[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->ssigma_tau_0[1][0])+(unsigned char)(cDiD->ssigma_tau_0[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->ssigma_tau_0[1][2])+(unsigned char)(cDiD->ssigma_tau_0[1][3])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->ssigma_tau_0[1]).Right(5));
      	oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->ssigma_tau_0[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->ssigma_tau_0[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sPsi_tau_int[3]!="")
      {
      	//184	��^2, ���������� �������� ��� ����������� C_m_int
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sPsi_tau_int[0]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(37)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sPsi_tau_int[1][0])+(unsigned char)(cDiD->sPsi_tau_int[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sPsi_tau_int[1][2])+(unsigned char)(cDiD->sPsi_tau_int[1][3])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sPsi_tau_int[1]).Right(4));
      	oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sPsi_tau_int[2]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSuperscript(True);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sPsi_tau_int[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sC_m_int[3]!="")
      {
      	//185	�, ������������ ����������� ��� ������� tau_m_int
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_m_int[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_m_int[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_m_int[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_m_int[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->stau_m_int[3]!="")
      {
      	//186	�, ��������� �������� ������������ ��������� � �������� ����������������� ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->stau_m_int[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->stau_m_int[1][0])+(unsigned char)(cDiD->stau_m_int[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->stau_m_int[1]).Right(6));
      	oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->stau_m_int[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->stau_m_int[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->ssigma_tau_int[3]!="")
      {
      	//187	��, ����������� ���������� ��� ��������� ������������ ��������� � �������� ������. ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->ssigma_tau_int[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->ssigma_tau_int[1][0])+(unsigned char)(cDiD->ssigma_tau_0[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->ssigma_tau_int[1][2])+(unsigned char)(cDiD->ssigma_tau_0[1][3])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->ssigma_tau_int[1]).Right(4));
      	oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->ssigma_tau_int[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->ssigma_tau_int[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sfi_tau_0[3]!="")
      {
      	//188	����������� ���������� � �������� ���������������� ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sfi_tau_0[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sfi_tau_0[1][0])+(unsigned char)(cDiD->sfi_tau_0[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sfi_tau_0[1][2])+(unsigned char)(cDiD->sfi_tau_0[1][3])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sfi_tau_0[1]).Right(1));
      	oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sfi_tau_0[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sfi_tau_0[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sfi_tau_int[3]!="")
      {
      	//189	����������� ���������� � �������� ����������������� ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sfi_tau_int[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sfi_tau_int[1][0])+(unsigned char)(cDiD->sfi_tau_int[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sfi_tau_int[1][2])+(unsigned char)(cDiD->sfi_tau_int[1][3])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sfi_tau_int[1]).Right(4));
      	oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sfi_tau_int[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sfi_tau_int[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      //190	����������� �������������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sK_int[0]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(66)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(6)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sK_int[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sK_int[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sK_int[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //191	%, ����������� ������ �� ������������ ����������� ������ ��� ��������� ������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sSESR[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sSESR[1]);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sSESR[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sSESR[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //192	%, ����������� ������������ ��� ��������� ������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sK_ng[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sK_ng[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sK_ng[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sK_ng[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //193	���� ��������-������������ ������ (����/���)
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sChRP[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sChRP[1]);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sChRP[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sChRP[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      if(cDiD->sK_stv[3]!="")
      {
      	//194	���������� ������� �������, �� �������� ����������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_stv[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_stv[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_stv[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_stv[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sDelta_f[3]!="")
      {
      	//195	���, ������ �� ������� ����� ��������� ������� � ��������� � ���� �������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_f[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_f[1][0])+(unsigned char)(cDiD->sDelta_f[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sDelta_f[1]).Right(1));
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_f[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_f[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sDelta_f_f[3]!="")
      {
      	//196	%, ��������� ���������� ������� � ������� �������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_f_f[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_f_f[1][0])+(unsigned char)(cDiD->sDelta_f_f[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sDelta_f_f[1]).Right(3));
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_f_f[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_f_f[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sC_Delta_f_priz[3]!="")
      {
      	//197	������������ �����������, ����������� �������������� ����������� ��������� �� ��������� ��� ��� ��������� ���������� ���� ��������������� ������� �� �������� Delta_f, � ����� ����������� ������ ������� �������������� ��� ����������������� ���������� ��� ��������� �����
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_Delta_f_priz[0]);
      	//��������� ����� "�����������"
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
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_Delta_f_priz[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_Delta_f_priz[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sC_f_priz[3]!="")
      {
      	//198	����������� C_f ��� ��������� �����
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_f_priz[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_f_priz[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_f_priz[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_f_priz[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sC_Delta_f[3]!="")
      {
      	//199	������������ �����������, ����������� �������������� ����������� ��������� �� ��������� ��� ��� ��������� ���������� ���� ��������������� ������� �� �������� Delta_f, � ����� ����������� ������ ������� �������������� ��� ����������������� ����������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_Delta_f[0]);
      	//��������� ����� "�����������"
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
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_Delta_f[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_Delta_f[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sC_f[3]!="")
      {
      	//200	����������� C_f
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_f[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_f[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_f[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_f[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sq[3]!="")
      {
      	//201	�����������, ����������� ����� �������, � ������� �������� ����� �������� ������� �� ������������ ��� �������������� ��� ����������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sq[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sq[1]);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sq[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sq[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sT_int_ChRP_priz[3]!="")
      {
      	//202	%, ����������������� �������������� �� ��������� ���� � ��� � ������ ����� ��� ��������� �����
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_int_ChRP_priz[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_int_ChRP_priz[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_int_ChRP_priz[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_int_ChRP_priz[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->salfa_tau_int_ChRP[3]!="")
      {
      	//203	����������� ������������ ������� � ��� � �������� ����������������� ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->salfa_tau_int_ChRP[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.InsertSymbol((256*(unsigned char)(cDiD->salfa_tau_int_ChRP[1][0])+(unsigned char)(cDiD->salfa_tau_int_ChRP[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->salfa_tau_int_ChRP[1][2])+(unsigned char)(cDiD->salfa_tau_int_ChRP[1][3])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->salfa_tau_int_ChRP[1]).Right(4));
      	oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->salfa_tau_int_ChRP[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->salfa_tau_int_ChRP[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sK_ng_ChRP_mes[3]!="")
      {
      	//204	%, ����������� ������������ �� ��������� ���� � ��� � ������ �����
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_ChRP_mes[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_ChRP_mes[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_ChRP_mes[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_ChRP_mes[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sT_ChRP[3]!="")
      {
      	//205	%, ��������� �������������� ����� �� ��������� ���� � ��� � ������ �����
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_ChRP[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_ChRP[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_ChRP[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_ChRP[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sSESR_ChRP[3]!="")
      {
      	//206	%, ����������� ������ �� ������������ ����������� ������ ��� ��� �� ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sSESR_ChRP[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sSESR_ChRP[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(4)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sSESR_ChRP[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sSESR_ChRP[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sK_ng_ChRP[3]!="")
      {
      	//207	%, ����������� ������������ �� ��������� ���� � ��� � ������� �� ���
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_ChRP[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_ChRP[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_ChRP[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_ChRP[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sPRP[3]!="")
      {
      	//208	���� ���������������-������������ ������ (����/���)
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sPRP[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sPRP[1]);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sPRP[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sPRP[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sDelta_h_rek_sleva[3]!="")
      {
      	//209	�, ������������� ������ ������ ��� ��� �����
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_h_rek_sleva[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_h_rek_sleva[1][0])+(unsigned char)(cDiD->sDelta_h_rek_sleva[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sDelta_h_rek_sleva[1]).Right(10));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(9)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_h_rek_sleva[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_h_rek_sleva[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sDelta_h_rek_sprava[3]!="")
      {
      	//210	�, ������������� ������ ������ ��� ��� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_h_rek_sprava[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_h_rek_sprava[1][0])+(unsigned char)(cDiD->sDelta_h_rek_sprava[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sDelta_h_rek_sprava[1]).Right(11));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(10)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_h_rek_sprava[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_h_rek_sprava[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sShema_PRP[3]!="")
      {
      	//211	����� ���������������-������������ ������ ("������������ (���.-���.)"/"�����-��-����� (���.-���.)")
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sShema_PRP[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sShema_PRP[1]);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sShema_PRP[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sShema_PRP[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sL_AVT2_dop[3]!="")
      {
      	//218	�, ����� ��������� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sL_AVT2_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sL_AVT2_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sL_AVT2_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sL_AVT2_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sh_sr_dop[3]!="")
      {
      	//219	�, ������� ������ ������ ���� ��� ������� ���� ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sh_sr_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sh_sr_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sh_sr_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sh_sr_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sGorn_dop[3]!="")
      {
      	//220	������ ������ ��� �������������� ������� ������ (���������/������/������������)
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sGorn_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sGorn_dop[1]);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sGorn_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sGorn_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sEta_AFT_dop[3]!="")
      {
      	//221	��, ��������� ������ �� ��������� � �������-�������� ������� ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sEta_AFT_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sEta_AFT_dop[1][0])+(unsigned char)(cDiD->sEta_AFT_dop[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sEta_AFT_dop[1]).Right(7));
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sEta_AFT_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sEta_AFT_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sW_0_dop[3]!="")
      {
      	//222	��, ���������� ������� � ��������� ������������ ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sW_0_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sW_0_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sW_0_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sW_0_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sR_pr_dop[3]!="")
      {
      	//223	��, ���������� �� ������������ ����������� �� ������� �������� �������������� �������� ��� ������� ��������� ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_pr_dop[0]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(54)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_pr_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_pr_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_pr_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sH_pr_dop[3]!="")
      {
      	//224	�, ������� � ����� ������������ ����������� �� ������� �������� �������������� �������� ��� ������� ��������� ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_pr_dop[0]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(56)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_pr_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_pr_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_pr_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sR1x_dop[3]!="")
      {
      	//225	��, ���������� �� ����� ������� ����� ��� ������� ��������� ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sR1x_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sR1x_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sR1x_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sR1x_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sR2x_dop[3]!="")
      {
      	//226	��, ���������� �� ������ ������� ����� ��� ������� ��������� ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sR2x_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sR2x_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sR2x_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sR2x_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sH1h_dop[3]!="")
      {
      	//227	�, �������� ������� ����� ������� ����� ��� ������� ��������� ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1h_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1h_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1h_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1h_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sH2h_dop[3]!="")
      {
      	//228	�, �������� ������� ������ ������� ����� ��� ������� ��������� ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2h_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2h_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2h_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2h_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sH_kr_h_dop[3]!="")
      {
      	//229	�, �������� ������� ����� � ����� �������� �������� ����������� �� �������� ������������ �������������� �������� ��� ������� ��������� ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_kr_h_dop[0]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(52)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_kr_h_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_kr_h_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_kr_h_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sR1_kas_dop[3]!="")
      {
      	//230	��, ���������� �� ����� ������� ����� ����������� � ������������ ������� ��� ������� ��������� ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sR1_kas_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sR1_kas_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sR1_kas_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sR1_kas_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sR2_kas_dop[3]!="")
      {
      	//231	��, ���������� �� ����� ������� ������ ����������� � ������������ ������� ��� ������� ��������� ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sR2_kas_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sR2_kas_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sR2_kas_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sR2_kas_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sH1_kas_dop[3]!="")
      {
      	//232	�, �������� ������� ����� ������� ����� ����������� � ������������ ������� ��� ������� ��������� ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_kas_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_kas_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_kas_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_kas_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sH2_kas_dop[3]!="")
      {
      	//233	�, �������� ������� ����� ������� ������ ����������� � ������������ ������� ��� ������� ��������� ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_kas_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_kas_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_kas_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_kas_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sR1_ellipse_dop[3]!="")
      {
      	//234	��, ���������� �� ����� ����������� ������ ����������� � ������������ ������� ��� ������� ��������� ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sR1_ellipse_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sR1_ellipse_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sR1_ellipse_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sR1_ellipse_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sR2_ellipse_dop[3]!="")
      {
      	//235	��, ���������� �� ����� ����������� ������� ����������� � ������������ ������� ��� ������� ��������� ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sR2_ellipse_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sR2_ellipse_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sR2_ellipse_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sR2_ellipse_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sH1_ellipse_dop[3]!="")
      {
      	//236	�, �������� ������� ����� ����������� ������ ����������� � ������������ ������� ��� ������� ��������� ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_ellipse_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_ellipse_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_ellipse_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_ellipse_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sH2_ellipse_dop[3]!="")
      {
      	//237	�, �������� ������� ����� ����������� ������� ����������� � ������������ ������� ��� ������� ��������� ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_ellipse_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_ellipse_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_ellipse_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_ellipse_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sR1_horda_H0_dop[3]!="")
      {
      	//238	��, ���������� �� ����� ������� �����, ����������� �� �������� H0, ��� ������� ��������� ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
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
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sR1_horda_H0_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sR1_horda_H0_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sR1_horda_H0_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sR2_horda_H0_dop[3]!="")
      {
      	//239	��, ���������� �� ������ ������� �����, ����������� �� �������� H0, ��� ������� ��������� ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
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
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sR2_horda_H0_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sR2_horda_H0_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sR2_horda_H0_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sH1_horda_H0_dop[3]!="")
      {
      	//240	�, �������� ������� ����� ������� �����, ����������� �� �������� H0, ��� ������� ��������� ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
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
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_horda_H0_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_horda_H0_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_horda_H0_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sH2_horda_H0_dop[3]!="")
      {
      	//241	�, �������� ������� ������ ������� �����, ����������� �� �������� H0, ��� ������� ��������� ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
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
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_horda_H0_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_horda_H0_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_horda_H0_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sR_cross_kas_dop[3]!="")
      {
      	//242	��, ���������� �� ����� ����������� ����������� � ����������� ������� ��� ������� ��������� ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_cross_kas_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_cross_kas_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_cross_kas_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_cross_kas_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sH_cross_kas_dop[3]!="")
      {
      	//243	�, �������� ������� ����� ����������� ����������� � ����������� ������� ��� ������� ��������� ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_cross_kas_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_cross_kas_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_cross_kas_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_cross_kas_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sProsvet_cross_kas_dop[3]!="")
      {
      	//244	�, ������� � ����� ����������� ����������� ��� ������� ��������� ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sProsvet_cross_kas_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sProsvet_cross_kas_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sProsvet_cross_kas_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sProsvet_cross_kas_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sR_kr_dop[3]!="")
      {
      	//245	��, ���������� �� ������������ ����������� ��� ������� ��������� ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_kr_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_kr_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_kr_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_kr_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sk_dop[3]!="")
      {
      	//246	������������� ���������� ������������ ����������� ��� ������� ��������� ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sk_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sk_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sk_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sk_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sH_kr_dop[3]!="")
      {
      	//247	�, ������� � ����� ������������ ����������� ��� ������� ��������� ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_kr_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_kr_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_kr_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_kr_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sH_0_dop[3]!="")
      {
      	//248	�, ����������� ������� � ����� ������������ ����������� ��� ������� ��������� ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_0_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_0_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_0_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_0_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sp_ot_g_kr_dop[3]!="")
      {
      	//249	������������� ������� � ����� ������������ ����������� ��� ������� ��������� ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sp_ot_g_kr_dop[0]);
      	//��������� ����� "�����������"
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
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sp_ot_g_kr_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sp_ot_g_kr_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sParametr_r_dop[3]!="")
      {
      	//250	��, �������� ����� ��� ������� ��������� ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sParametr_r_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sParametr_r_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sParametr_r_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sParametr_r_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sDelta_y_dop[3]!="")
      {
      	//251	�, ������ �������� ���������������� ����� ��� ������� ��������� ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_y_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_y_dop[1][0])+(unsigned char)(cDiD->sDelta_y_dop[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sDelta_y_dop[1]).Right(8));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(7)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_y_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_y_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sb_dop[3]!="")
      {
      	//252	k�, ������ ���������������� ����� ��� ������� ��������� ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sb_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sb_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sb_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sb_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->smu_dop[3]!="")
      {
      	//253	��������, ��������������� �������� ���������������� ����� ��� ������� ��������� ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->smu_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->smu_dop[1][0])+(unsigned char)(cDiD->smu_dop[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->smu_dop[1]).Right(7));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->smu_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->smu_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sV_difr_sr_dop[3]!="")
      {
      	//254	��, ������������� ���������� ��� ������� ��������� ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_difr_sr_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_difr_sr_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_difr_sr_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_difr_sr_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sV_50_proc_dop[3]!="")
      {
      	//255	��, ������� ���������� ������� ��-�� �������� ����� �� ������ � ������������ ������� ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_50_proc_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_50_proc_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(4)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_50_proc_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_50_proc_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sP_pm_dop[3]!="")
      {
      	//256	���, ������� ������� ������� �� ����� ��������� ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sP_pm_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sP_pm_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sP_pm_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sP_pm_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sV_0_min_d_dop[3]!="")
      {
      	//257	��, ���������� ���������� ��������� ���������� �� ������ ��� ����� ���������� ���������� ������, ������� ����������, � ������ �������� �� ������� ��������� ������������ ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_0_min_d_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_0_min_d_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_0_min_d_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_0_min_d_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sV_0_min_0_dop[3]!="")
      {
      	//258	��, ���������� ���������� ��������� ���������� �� ������������ � ���. ��� ����� ���������� ���������� ������, ������� ����������, � ������ �������� �� ������� ��������� ������������ ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_0_min_0_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_0_min_0_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_0_min_0_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_0_min_0_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sZ_por_dop_dop[3]!="")
      {
      	//259	��, ��������� ��������� (������/��� ���, ��� ���������� ���������� ������ �� 3 ��) ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_por_dop_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_por_dop_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_por_dop_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_por_dop_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sZ_sosedn_kanal_dop[3]!="")
      {
      	//260	��, ���������� ���������� (������/��� ���, ��� ���������� ���������� ������ �� 3 ��) ���������� ��������� (������ �� ��������� ������/��� ���) ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
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
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_sosedn_kanal_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_sosedn_kanal_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_sosedn_kanal_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sDelta_V_sosedn_kanal_dop[3]!="")
      {
      	//261	��, ���������� ���������� ������ ��-�� ������� ����� �� ��������� ������ ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_sosedn_kanal_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_V_sosedn_kanal_dop[1][0])+(unsigned char)(cDiD->sDelta_V_sosedn_kanal_dop[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sDelta_V_sosedn_kanal_dop[1]).Right(7));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(6)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_sosedn_kanal_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_sosedn_kanal_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sD_p_0_dop[3]!="")
      {
      	//262	��, ����������� ��������������� ������ ��� ���������� ��������� �������������� �������� �������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sD_p_0_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sD_p_0_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sD_p_0_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sD_p_0_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sQ_p_cochannel_dop[3]!="")
      {
      	//263	��, �����������, ����������� ������ �������������������� ��������� �������������� ������� ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sQ_p_cochannel_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sQ_p_cochannel_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sQ_p_cochannel_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sQ_p_cochannel_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sD_p_dop[3]!="")
      {
      	//264	��, ����������� ��������������� ������ � �������� ����������������� ��������� ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sD_p_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sD_p_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sD_p_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sD_p_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sD_p_dojd_dop[3]!="")
      {
      	//265	��, ����������� ��������������� ������, ������������� �������� ������ ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sD_p_dojd_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sD_p_dojd_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sD_p_dojd_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sD_p_dojd_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sZ_polyariz_cochannel_dop[3]!="")
      {
      	//266	��, ���������� ���������� (�������������������� ������/��� ���) ���������� ��������� (������/��� ���, ��� ���������� ���������� ������ �� 3 ��), ��� ����� � �������� ���� (� �������� �0 �������������������� ������ �� �����������) ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
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
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_polyariz_cochannel_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_polyariz_cochannel_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_polyariz_cochannel_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sZ_polyariz_cochannel_dojd_dop[3]!="")
      {
      	//267	��, ���������� ���������� (�������������������� ������/��� ���) ���������� ��������� (������/��� ���, ��� ���������� ���������� ������ �� 3 ��), ��� ����� � �������� �� (� �������� �0 �������������������� ������ �� �����������) ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
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
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_polyariz_cochannel_dojd_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_polyariz_cochannel_dojd_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_polyariz_cochannel_dojd_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sDelta_V_polyariz_cochannel_dop[3]!="")
      {
      	//268	��, ���������� ���������� ������ ��-�� ������� �������������������� ����� ��� ����� � �������� ���� (� �������� �0 �������������������� ������ �� �����������) ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_polyariz_cochannel_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_V_polyariz_cochannel_dop[1][0])+(unsigned char)(cDiD->sDelta_V_polyariz_cochannel_dop[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sDelta_V_polyariz_cochannel_dop[1]).Right(11));
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_polyariz_cochannel_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_polyariz_cochannel_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sDelta_V_polyariz_cochannel_dojd_dop[3]!="")
      {
      	//269	��, ���������� ���������� ������ ��-�� ������� �������������������� ����� ��� ����� � �������� �� (� �������� �0 �������������������� ������ �� �����������) ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_polyariz_cochannel_dojd_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_V_polyariz_cochannel_dojd_dop[1][0])+(unsigned char)(cDiD->sDelta_V_polyariz_cochannel_dojd_dop[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sDelta_V_polyariz_cochannel_dojd_dop[1]).Right(9));
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_polyariz_cochannel_dojd_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_polyariz_cochannel_dojd_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sZ_por_dopoln[3]!="")
      {
      	//270	��, ��������� (��������� ������ ��������� ����������� � ���������������/��� ���) ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_por_dopoln[0]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(49)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sZ_por_dopoln[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_por_dopoln[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_por_dopoln[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sZ_obr_uzl_dop[3]!="")
      {
      	//271	��, ���������� ���������� (��������� ������ ��������� ����������� � ���������������/��� ���) ���������� ��������� (������/��� ���, ��� ���������� ���������� ������ �� 3 ��) ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
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
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sZ_obr_uzl_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_obr_uzl_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sZ_obr_uzl_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sDelta_V_obr_uzl_dop[3]!="")
      {
      	//272	��, ���������� ���������� ������ ��-�� ������� ����� � ��������� ����������� � ��������������� ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_obr_uzl_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_V_obr_uzl_dop[1][0])+(unsigned char)(cDiD->sDelta_V_obr_uzl_dop[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
         oSel.SetText((cDiD->sDelta_V_obr_uzl_dop[1]).Right(13));
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_obr_uzl_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_obr_uzl_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sDelta_V_degr_int_dop[3]!="")
      {
      	//273	��, ���������� ���������� ������ ��-�� ������� ����� �� T��� ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_degr_int_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_V_degr_int_dop[1][0])+(unsigned char)(cDiD->sDelta_V_degr_int_dop[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
         oSel.SetText((cDiD->sDelta_V_degr_int_dop[1]).Right(14));
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_degr_int_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_degr_int_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sDelta_V_degr_subrefr_dop[3]!="")
      {
      	//274	��, ���������� ���������� ������ ��-�� ������� ����� �� T0 ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_degr_subrefr_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_V_degr_subrefr_dop[1][0])+(unsigned char)(cDiD->sDelta_V_degr_subrefr_dop[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
         oSel.SetText((cDiD->sDelta_V_degr_subrefr_dop[1]).Right(15));
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_degr_subrefr_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_degr_subrefr_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sDelta_V_degr_dojd_dop[3]!="")
      {
      	//275	��, ���������� ���������� ������ ��-�� ������� ����� �� T� ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_degr_dojd_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_V_degr_dojd_dop[1][0])+(unsigned char)(cDiD->sDelta_V_degr_dojd_dop[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
         oSel.SetText((cDiD->sDelta_V_degr_dojd_dop[1]).Right(15));
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_degr_dojd_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_degr_dojd_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sV_min_int_dop[3]!="")
      {
      	//276	��, ���������� ���������� ��������� ������������������ ���������� c ������ ���������� ���������� ������, ������� ���������� � �������� �� ������� ��������� ������������ ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_int_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sV_min_int_dop[1]);
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_int_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_int_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sV_min_subrefr_dop[3]!="")
      {
      	//277	��, ���������� ���������� ��������� ����������������� ���������� c ������ ���������� ���������� ������, ������� ���������� � �������� �� ������� ��������� ������������ ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_subrefr_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sV_min_subrefr_dop[1]);
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_subrefr_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_subrefr_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sV_min_eff_dop[3]!="")
      {
      	//278	��, ���������� ���������� ����������� ��������� ������������������ ���������� ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_eff_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sV_min_eff_dop[1]);
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_eff_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_eff_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sR_otr_dop[3]!="")
      {
      	//279	��, ���������� �� ����� ��������� ��� �������������� ������� ������ ��� ������� ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_otr_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sR_otr_dop[1]);
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_otr_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_otr_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sH_otr_dop[3]!="")
      {
      	//280	�, ������� � ����� ��������� ��� �������������� ������� ������ ��� ������� ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_otr_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sH_otr_dop[1]);
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_otr_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_otr_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sH_0_otr_dop[3]!="")
      {
      	//281	�, ����������� ������� � ����� ��������� ��� �������������� ������� ������ ��� ������� ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_0_otr_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sH_0_otr_dop[1]);
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_0_otr_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_0_otr_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sp_ot_g_sr_otr_dop[3]!="")
      {
      	//282	������������� ������� � ����� ��������� ��� �������������� ������� ������ ��� ������� ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sp_ot_g_sr_otr_dop[0]);
      	//��������� ����� "�����������"
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
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sp_ot_g_sr_otr_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sp_ot_g_sr_otr_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sX_prodoln_dop[3]!="")
      {
      	//283	��, ������������� ����������� ������� �������, ��������� ��� ������������������ ��������� ������, ��� �������������� ������� ������ ��� ������� ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sX_prodoln_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sX_prodoln_dop[1]);
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sX_prodoln_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sX_prodoln_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sDelta_h_max_dop[3]!="")
      {
      	//284	�, ������������ ���������� ������� ����������� ������� ������� �� ���������������� �����������, ��� ������� ��� ����� �������, ��� ������� �����������������, �������������� ������� ������ ��� ������� ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_h_max_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_h_max_dop[1][0])+(unsigned char)(cDiD->sDelta_h_max_dop[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
         oSel.SetText((cDiD->sDelta_h_max_dop[1]).Right(15));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(14)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_h_max_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_h_max_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sA_otr_dop[3]!="")
      {
      	//285	�������� � ��� ����� ��������� ��� ������� ��������� ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sA_otr_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sA_otr_dop[1]);
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sA_otr_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sA_otr_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sF_ot_p_g_A_dop[3]!="")
      {
      	//286	�������� F(p(g),A) ��� ����� ��������� ��� ������� ��������� ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sF_ot_p_g_A_dop[0]);
      	//��������� ����� "�����������"
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
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sF_ot_p_g_A_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sF_ot_p_g_A_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sQ_dop[3]!="")
      {
      	//287	�������� Q ��� ������� ��������� ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sQ_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sQ_dop[1]);
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sQ_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sQ_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sT_int_dop[3]!="")
      {
      	//288	%, ��������������, ������������� ������������������ ��������� ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_int_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sT_int_dop[1]);
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_int_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_int_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sfi_tau_int_dop[3]!="")
      {
      	//289	����������� ���������� � �������� ����������������� ��������� ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sfi_tau_int_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sfi_tau_int_dop[1][0])+(unsigned char)(cDiD->sfi_tau_int_dop[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sfi_tau_int_dop[1][2])+(unsigned char)(cDiD->sfi_tau_int_dop[1][3])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sfi_tau_int_dop[1]).Right(8));
      	oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sfi_tau_int_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sfi_tau_int_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sfi_tau_0_dop[3]!="")
      {
      	//290	����������� ���������� � �������� ���������������� ��������� ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sfi_tau_0_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sfi_tau_0_dop[1][0])+(unsigned char)(cDiD->sfi_tau_0_dop[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sfi_tau_0_dop[1][2])+(unsigned char)(cDiD->sfi_tau_0_dop[1][3])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sfi_tau_0_dop[1]).Right(5));
      	oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sfi_tau_0_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sfi_tau_0_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sK_int_dop[3]!="")
      {
      	//291	����������� ������������� ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_int_dop[0]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(54)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(6)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_int_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_int_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_int_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sSESR_dop[3]!="")
      {
      	//292	%, ����������� ������ �� ������������ ����������� ������ ��� ��������� ������ ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sSESR_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sSESR_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(4)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sSESR_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sSESR_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sK_ng_dop[3]!="")
      {
      	//293	%, ����������� ������������ ��� ��������� ������ ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sPsi_tau_0_dop[3]!="")
      {
      	//294	��^2, ���������� �������� ��� ����������� C_m_0 ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sPsi_tau_0_dop[0]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(28)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sPsi_tau_0_dop[1][0])+(unsigned char)(cDiD->sPsi_tau_0_dop[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sPsi_tau_0_dop[1][2])+(unsigned char)(cDiD->sPsi_tau_0_dop[1][3])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sPsi_tau_0_dop[1]).Right(9));
      	oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sPsi_tau_0_dop[2]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSuperscript(True);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sPsi_tau_0_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sC_m_0_dop[3]!="")
      {
      	//295	�, ������������ ����������� ��� ������� tau_m_0 ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_m_0_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_m_0_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_m_0_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_m_0_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->stau_m_0_dop[3]!="")
      {
      	//296	�, ��������� �������� ������������ ��������� � �������� ���������������� ��������� ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->stau_m_0_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->stau_m_0_dop[1][0])+(unsigned char)(cDiD->stau_m_0_dop[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->stau_m_0_dop[1]).Right(11));
      	oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->stau_m_0_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->stau_m_0_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->ssigma_tau_0_dop[3]!="")
      {
      	//297	��, ����������� ���������� ��� ��������� ������������ ��������� � �������� �������. ��������� ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->ssigma_tau_0_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->ssigma_tau_0_dop[1][0])+(unsigned char)(cDiD->ssigma_tau_0_dop[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->ssigma_tau_0_dop[1][2])+(unsigned char)(cDiD->ssigma_tau_0_dop[1][3])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->ssigma_tau_0_dop[1]).Right(9));
      	oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->ssigma_tau_0_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->ssigma_tau_0_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sPsi_tau_int[3]!="")
      {
      	//298	��^2, ���������� �������� ��� ����������� C_m_int ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sPsi_tau_int_dop[0]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(33)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sPsi_tau_int_dop[1][0])+(unsigned char)(cDiD->sPsi_tau_int_dop[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sPsi_tau_int_dop[1][2])+(unsigned char)(cDiD->sPsi_tau_int_dop[1][3])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sPsi_tau_int_dop[1]).Right(8));
      	oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sPsi_tau_int_dop[2]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSuperscript(True);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sPsi_tau_int_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sC_m_int_dop[3]!="")
      {
      	//299	�, ������������ ����������� ��� ������� tau_m_int ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_m_int_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_m_int_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_m_int_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_m_int_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->stau_m_int_dop[3]!="")
      {
      	//300	�, ��������� �������� ������������ ��������� � �������� ����������������� ��������� ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->stau_m_int_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->stau_m_int_dop[1][0])+(unsigned char)(cDiD->stau_m_int_dop[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->stau_m_int_dop[1]).Right(10));
      	oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->stau_m_int_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->stau_m_int_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->ssigma_tau_int_dop[3]!="")
      {
      	//301	��, ����������� ���������� ��� ��������� ������������ ��������� � �������� ������. ��������� ��� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->ssigma_tau_int_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->ssigma_tau_int_dop[1][0])+(unsigned char)(cDiD->ssigma_tau_int_dop[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->ssigma_tau_int_dop[1][2])+(unsigned char)(cDiD->ssigma_tau_int_dop[1][3])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->ssigma_tau_int_dop[1]).Right(8));
      	oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->ssigma_tau_int_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->ssigma_tau_int_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sC_Delta_h[3]!="")
      {
      	//302	������������ �����������, ����������� �������������� ����������� ��������� ��� ���������������� ���������� ������ (��� �������� �������, ���� ������ �����������������)
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_Delta_h[0]);
      	//��������� ����� "�����������"
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
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_Delta_h[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_Delta_h[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sC_Delta_h_dop[3]!="")
      {
      	//303	������������ �����������, ����������� �������������� ����������� ��������� ��� ���������������� ���������� ������ ��� �������������� ������� ��� ����������������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_Delta_h_dop[0]);
      	//��������� ����� "�����������"
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
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_Delta_h_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sC_Delta_h_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sI_PRP_peres[3]!="")
      {
      	//304	��, ������������� ��� ��� ������� ��� ��� �� ��������� � ���������� ������ ��� ������������ ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sI_PRP_peres[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sI_PRP_peres[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sI_PRP_peres[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sI_PRP_peres[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sT_int_PRP[3]!="")
      {
      	//305	%, ��������������, ������������� ������������������ ��������� ��� ���
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_int_PRP[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_int_PRP[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_int_PRP[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_int_PRP[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sfi_tau_int_PRP[3]!="")
      {
      	//306	����������� ���������� � �������� ����������������� ��������� ��� ���
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sfi_tau_int_PRP[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sfi_tau_int_PRP[1][0])+(unsigned char)(cDiD->sfi_tau_int_PRP[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sfi_tau_int_PRP[1][2])+(unsigned char)(cDiD->sfi_tau_int_PRP[1][3])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sfi_tau_int_PRP[1]).Right(7));
      	oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sfi_tau_int_PRP[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sfi_tau_int_PRP[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sSESR_PRP[3]!="")
      {
      	//307	%, ����������� ������ �� ������������ ����������� ������ ��� ���
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sSESR_PRP[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sSESR_PRP[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(4)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sSESR_PRP[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sSESR_PRP[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sK_ng_PRP[3]!="")
      {
      	//308	%, ����������� ������������  ��� ���
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_PRP[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_PRP[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_PRP[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_PRP[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sSESR_norm[3]!="")
      {
      	//309	%, ����� �� ����������� ������ �� ������������ ����������� ������ ��� ��������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sSESR_norm[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sSESR_norm[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(4)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sSESR_norm[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sSESR_norm[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sK_ng_norm[3]!="")
      {
      	//310	%, ����� �� ����������� ������������ ��� ��������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_norm[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_norm[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_norm[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_norm[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }
   }
   //������� �����
   else
   {


      //9	��� ������������
      //���������� ������
      oCell = oTable.Cell(N_strok+2,3);
   	oCell.Select();
   	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdExtend)));
   	oCells=oSel.GetCells();
   	oCells.Merge();
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sTyp_obor[0]);
      //��������� ����� "�����������", "�������.", "��������"
      oCell = oTable.Cell(N_strok+2,3);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sTyp_obor[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //10	���������� ��� (������������� �����, ��)
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sNaznachenie[0]);
      //��������� ����� "�����������", "�������.", "��������"
      oCell = oTable.Cell(N_strok+2,3);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sNaznachenie[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //12	�������� ������(�� ���������, �� ��������������, �� ������ �������)
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sHarakter_trassy[0]);
      //��������� ����� "�����������", "�������.", "��������"
      oCell = oTable.Cell(N_strok+2,3);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sHarakter_trassy[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      if(cDiD->sPol[3]!="")
      {
      	//13	�����������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oRan = oCell.GetRange();
      	oRan.SetText(cDiD->sPol[0]);
      	//��������� ����� "�����������", "�������.", "��������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oRan = oCell.GetRange();
      	oRan.SetText(cDiD->sPol[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      //14	��� ������� ���������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sTyp_sys[0]);
      //��������� ����� "�����������", "�������.", "��������"
      oCell = oTable.Cell(N_strok+2,3);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sTyp_sys[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //15	��, ������������� ���������
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sR[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sR[1]);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sR[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sR[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //16	���, ��������� ��������
      //��������� ������
      oCell.Select();
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sf[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sf[1]);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sf[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oRan = oCell.GetRange();
      oRan.SetText(cDiD->sf[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //18	����/�, �������� �������� ��������� ������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sC[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sC[1]);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sC[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sC[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //19	���, �������� �����������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sP_pd[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sP_pd[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
   	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sP_pd[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sP_pd[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


       //23	���, ��������� ���������������� ��������� ��� BER
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sP_pm_por[0]);
      //��������� ���������
      	oParFormat=oSel.GetParagraphFormat();
      	CTabStops oTabStops;
      	oTabStops=oParFormat.get_TabStops();
      	oTabStops.Add(80.2*MM2PH, COleVariant(short(wdAlignTabLeft)), COleVariant(short(wdTabLeaderSpaces)));
      	oParFormat.put_TabStops(oTabStops);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.SetText(cDiD->sBER[3]);//22	����������� ������ �� �����, � ����������� �� �������� �������� ��������� ������
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sP_pm_por[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
   	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sP_pm_por[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sP_pm_por[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //30	���, ����������� �������� �������� ������� �����
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.SetText(cDiD->sG_pd[0]);
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.InsertSymbol(216, COleVariant("Arial"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      oSel.SetText(cDiD->sD_pd[3]);//29	�, ������� �������� ������� �����
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sG_pd[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
   	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sG_pd[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sG_pd[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //32	���, ����������� �������� �������� ������� ������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.SetText(cDiD->sG_pm[0]);
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.InsertSymbol(216, COleVariant("Arial"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      oSel.SetText(cDiD->sD_pm[3]);//31	�, ������� �������� ������� ������
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sG_pm[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
   	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sG_pm[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sG_pm[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      if(cDiD->sG1_dop[3]!="")
      {
      	//215	���, ����������� �������� �������������� ������� ����� (����������� ��������/� ������ �����������)
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.SetText(cDiD->sG1_dop[0]);
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.InsertSymbol(216, COleVariant("Arial"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText(cDiD->sD1_dop[3]);//214	�, ������� �������������� ������� �����
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sG1_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sG1_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sG1_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sG2_dop[3]!="")
      {
      	//217	���, ����������� �������� �������������� ������� ������ (����������� ��������/� ������ �����������)
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.SetText(cDiD->sG2_dop[0]);
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.InsertSymbol(216, COleVariant("Arial"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText(cDiD->sD2_dop[3]);//216	�, ������� �������������� ������� ������
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sG2_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sG2_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sG2_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      //39	�, ������ ������ �������� �������� ������� �����
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sH1[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sH1[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
   	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sH1[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sH1[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //40	�, ������ ������ �������� �������� ������� ������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sH2[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sH2[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
   	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sH2[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sH2[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      if(cDiD->sH1_dop[3]!="")
      {
      	//212	�, ������ ������ �������� �������������� ������� �����
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH1_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sH2_dop[3]!="")
      {
      	//213	�, ������ ������ �������� �������������� ������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_dop[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_dop[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_dop[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH2_dop[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      //43	1/�, ������� �������� ��������� ����. ������������� �������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sg_sr[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sg_sr[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
   	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sg_sr[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sg_sr[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //44	1/�, ����������� ���������� ��������� ����. ������������� �������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sSigma[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.InsertSymbol((256*(unsigned char)(cDiD->sSigma[1][0])+(unsigned char)(cDiD->sSigma[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sSigma[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sSigma[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      if(cDiD->sSigma_ot_R[3]!="")
      {
      	//45	1/�, ��� ��������� ����. ������������� ������� � ����������� �� ����������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sSigma_ot_R[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sSigma_ot_R[1][0])+(unsigned char)(cDiD->sSigma_ot_R[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sSigma_ot_R[1]).Right(3));
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sSigma_ot_R[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sSigma_ot_R[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      //46	��, ��������� ������ �� ��������� � �������-�������� �������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sEta_AFT[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.InsertSymbol((256*(unsigned char)(cDiD->sEta_AFT[1][0])+(unsigned char)(cDiD->sEta_AFT[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      oSel.SetText((cDiD->sEta_AFT[1]).Right(3));
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
   	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sEta_AFT[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sEta_AFT[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //47	��, ���������� ������� � ��������� ������������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sW_0[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sW_0[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
   	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sW_0[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sW_0[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      if(cDiD->sV_difr_sr[3]!="")
      {
      	//48	��, ������� ���������� �� ���� ��������� (��� ������� ���������)
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_difr_sr[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_difr_sr[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_difr_sr[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_difr_sr[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sV_v_gazah[3]!="")
      {
      	//50	��, ������� ���������� ������� � ����� ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_v_gazah[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_v_gazah[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_v_gazah[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_v_gazah[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sV_50_proc[3]!="")
      {
      	//51	��, ������� ���������� ������� ��-�� �������� ����� �� ������ � ������������ �������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_50_proc[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_50_proc[1]);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_50_proc[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_50_proc[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      //52	���, ������� ������� ������� �� ����� ���������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sP_pm[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sP_pm[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
   	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sP_pm[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sP_pm[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      if(cDiD->sDelta_V_tip[3]!="")
      {
      	//53	��, �������� �� Vmin ��� ������� ���������� ����������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_tip[0]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(13)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(3)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_V_tip[1][0])+(unsigned char)(cDiD->sDelta_V_tip[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sDelta_V_tip[1]).Right(5));
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
   		oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_tip[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_tip[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sDelta_V_degr_int[3]!="")
      {
      	//81	��, ���������� ���������� ������ ��-�� ������� ����� �� T���
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_degr_int[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_V_degr_int[1][0])+(unsigned char)(cDiD->sDelta_V_degr_int[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
         oSel.SetText((cDiD->sDelta_V_degr_int[1]).Right(10));
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_degr_int[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_degr_int[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sDelta_V_degr_subrefr[3]!="")
      {
      	//82	��, ���������� ���������� ������ ��-�� ������� ����� �� T0
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_degr_subrefr[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_V_degr_subrefr[1][0])+(unsigned char)(cDiD->sDelta_V_degr_subrefr[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
         oSel.SetText((cDiD->sDelta_V_degr_subrefr[1]).Right(11));
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_degr_subrefr[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_degr_subrefr[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sDelta_V_degr_dojd[3]!="")
      {
      	//83	��, ���������� ���������� ������ ��-�� ������� ����� �� T�
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_degr_dojd[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_V_degr_dojd[1][0])+(unsigned char)(cDiD->sDelta_V_degr_dojd[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
         oSel.SetText((cDiD->sDelta_V_degr_dojd[1]).Right(11));
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_degr_dojd[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_V_degr_dojd[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sV_min_int[3]!="")
      {
      	//84	��, ���������� ���������� ��������� ������������������ ���������� c ������ ���������� ���������� ������, ������� ���������� � �������� �� ������� ��������� ������������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_int[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sV_min_int[1]);
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_int[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_int[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sV_min_subrefr[3]!="")
      {
      	//85	��, ���������� ���������� ��������� ����������������� ���������� c ������ ���������� ���������� ������, ������� ���������� � �������� �� ������� ��������� ������������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_subrefr[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sV_min_subrefr[1]);
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_subrefr[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_subrefr[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sV_min_d[3]!="")
      {
      	//86	��, ���������� ���������� ��������� ���������� � ����� c ������ ���������� ���������� ������, ������� ���������� � �������� �� ������� ��������� ������������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_d[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sV_min_d[1]);
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_d[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_d[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sh_s[3]!="")
      {
      	//88	��, ������ ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sh_s[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sh_s[1]);
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sh_s[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sh_s[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sDelta_f_s[3]!="")
      {
      	//89	���, ������ ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_f_s[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_f_s[1][0])+(unsigned char)(cDiD->sDelta_f_s[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
         oSel.SetText((cDiD->sDelta_f_s[1]).Right(2));
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_f_s[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_f_s[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sS_s[3]!="")
      {
      	//90	���, ������� ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sS_s[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sS_s[1]);
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sS_s[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sS_s[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sI_ekv[3]!="")
      {
      	//91	��, ������� �� ���� �����������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sI_ekv[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sI_ekv[1]);
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sI_ekv[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sI_ekv[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sV_min_eff_pred_ekv[3]!="")
      {
      	//93	��, ��������� ����������� ���������� ���������� ����������� ��������� ������������������ ���������� � ������ �����������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_eff_pred_ekv[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sV_min_eff_pred_ekv[1]);
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_eff_pred_ekv[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_eff_pred_ekv[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sV_min_eff[3]!="")
      {
      	//94	��, ���������� ���������� ����������� ��������� ������������������ ����������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_eff[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
         oSel.SetText(cDiD->sV_min_eff[1]);
         oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_eff[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sV_min_eff[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      //97	��, ���������� �� ������������ ����������� ��� ���������� ���������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sR_kr[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sR_kr[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sR_kr[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sR_kr[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //99	�, ������� � ����� ������������ ����������� ��� ���������� ���������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sH_kr[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sH_kr[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sH_kr[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sH_kr[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //100	�, ����������� ������� � ����� ������������ ����������� ��� ���������� ���������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sH_0[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sH_0[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sH_0[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sH_0[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //101	��, �������� ����� ��� ���������� ���������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sParametr_r[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sParametr_r[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sParametr_r[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sParametr_r[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //105	��������, ��������������� �������� ���������������� ����� ��� ���������� ���������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->smu_0[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.InsertSymbol((256*(unsigned char)(cDiD->smu_0[1][0])+(unsigned char)(cDiD->smu_0[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      oSel.SetText((cDiD->smu_0[1]).Right(1));
      oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->smu_0[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->smu_0[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //106	�������� � ��� ������������ ����������� ��� ���������� ���������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sA_kr[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sA_kr[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sA_kr[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sA_kr[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //114	������������� ������� � ����� ������������ ����������� ��� ������� ���������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sp_ot_g_sr_kr[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sp_ot_g_sr_kr[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(3)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sp_ot_g_sr_kr[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sp_ot_g_sr_kr[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //119	1/�, ��������� �������� ������������ ��������� ��������������� ������������� ������� ��� V����=V���.
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sg_0[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sg_0[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sg_0[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sg_0[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //145	������������� ������� � ����� ������������ ����������� ��� ��������� ���������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sp_ot_g_0_kr[0]);
      //��������� ����� "�����������"
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
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sp_ot_g_0_kr[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sp_ot_g_0_kr[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //149	��������, ��������������� �������� ���������������� ����� ��� ��������� ���������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->smu_3[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.InsertSymbol((256*(unsigned char)(cDiD->smu_3[1][0])+(unsigned char)(cDiD->smu_3[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      oSel.SetText((cDiD->smu_3[1]).Right(4));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->smu_3[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->smu_3[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //151	�������� ���
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sPsi[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.InsertSymbol((256*(unsigned char)(cDiD->sPsi[1][0])+(unsigned char)(cDiD->sPsi[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sPsi[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sPsi[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //152	%, ��������������, ������������� �������������� ���������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sT_0[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sT_0[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sT_0[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sT_0[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      if(cDiD->sR_otr_2[3]!="")
      {
      	//159	��, ���������� �� ����� ��������� ��� �������� ������ ��� ������� ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_otr_2[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_otr_2[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_otr_2[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sR_otr_2[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sH_otr_2[3]!="")
      {
      	//160	�, ������� � ����� ��������� ��� �������� ������ ��� ������� ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_otr_2[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_otr_2[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_otr_2[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sH_otr_2[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sp_ot_g_sr_otr[3]!="")
      {
      	//162	������������� ������� � ����� ��������� ��� �������� ������ ��� ������� ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sp_ot_g_sr_otr[0]);
      	//��������� ����� "�����������"
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
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sp_ot_g_sr_otr[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sp_ot_g_sr_otr[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sA_otr[3]!="")
      {
      	//165	�������� � ��� ����� ��������� ��� ������� ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sA_otr[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sA_otr[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sA_otr[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sA_otr[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sF_ot_p_g_A[3]!="")
      {
      	//166	�������� F(p(g),A) ��� ����� ��������� ��� ������� ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sF_ot_p_g_A[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sF_ot_p_g_A[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(5)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sF_ot_p_g_A[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sF_ot_p_g_A[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sK_vp[3]!="")
      {
      	//167	%, ����������� ������ �����������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_vp[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_vp[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_vp[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_vp[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      //168	�������� Q ��� ������� ���������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sQ[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sQ[1]);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sQ[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sQ[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //169	%, ��������, ����������� ����������� ������������� ������������ ���������, ������������� ����������� ��������� �� �������� ��������������� ���������� � ��������� ��������������� ������������� ������� Delta_epsilon
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sT_Delta_eps[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText((cDiD->sT_Delta_eps[1]).Left(2));
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdMove)));
      oSel.InsertSymbol((256*(unsigned char)(cDiD->sT_Delta_eps[1][2])+(unsigned char)(cDiD->sT_Delta_eps[1][3])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      oSel.InsertSymbol((256*(unsigned char)(cDiD->sT_Delta_eps[1][4])+(unsigned char)(cDiD->sT_Delta_eps[1][5])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      oSel.SetText((cDiD->sT_Delta_eps[1]).Right(1));
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sT_Delta_eps[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sT_Delta_eps[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //173	%, ��������������, ������������� ������������������ ���������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sT_int[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sT_int[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sT_int[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sT_int[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      if(cDiD->sT_d_m[3]!="")
      {
      	//177	%, ��������������, ������������� �������� ������� � ��������� �����
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_d_m[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_d_m[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_d_m[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_d_m[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sQ_d[3]!="")
      {
      	//178	����������� ��������� �������� ���������� ������ � �������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sQ_d[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sQ_d[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sQ_d[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sQ_d[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sT_d_g[3]!="")
      {
      	//179	%, ��������������, ������������� �������� ������� � ������� �� ���
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_d_g[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_d_g[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_d_g[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sT_d_g[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sfi_tau_0[3]!="")
      {
      	//188	����������� ���������� � �������� ���������������� ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sfi_tau_0[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sfi_tau_0[1][0])+(unsigned char)(cDiD->sfi_tau_0[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sfi_tau_0[1][2])+(unsigned char)(cDiD->sfi_tau_0[1][3])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sfi_tau_0[1]).Right(1));
      	oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sfi_tau_0[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sfi_tau_0[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sfi_tau_int[3]!="")
      {
      	//189	����������� ���������� � �������� ����������������� ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sfi_tau_int[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sfi_tau_int[1][0])+(unsigned char)(cDiD->sfi_tau_int[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sfi_tau_int[1][2])+(unsigned char)(cDiD->sfi_tau_int[1][3])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sfi_tau_int[1]).Right(4));
      	oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sfi_tau_int[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sfi_tau_int[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      //190	����������� �������������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sK_int[0]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(66)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(6)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sK_int[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sK_int[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sK_int[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //191	%, ����������� ������ �� ������������ ����������� ������ ��� ��������� ������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sSESR[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sSESR[1]);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sSESR[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sSESR[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      //192	%, ����������� ������������ ��� ��������� ������
      //��������� ������
      oSel.SelectRow();
      oSel.InsertRowsBelow(COleVariant(short(1)));
      //��������� ����� "������������ ���������"
      oCell = oTable.Cell(N_strok+2,2);
      oCell.Select();
      oSel.SetText(cDiD->sK_ng[0]);
      //��������� ����� "�����������"
      oCell = oTable.Cell(N_strok+2,3);
      oCell.Select();
      oSel.SetText(cDiD->sK_ng[1]);
      oSel.Collapse(COleVariant(short(wdCollapseStart)));
      oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      oFont.SetSubscript(True);
      //��������� ����� "�������."
      oCell = oTable.Cell(N_strok+2,4);
      oCell.Select();
      oSel.SetText(cDiD->sK_ng[2]);
      //��������� ����� "��������"
      oCell = oTable.Cell(N_strok+2,5);
      oCell.Select();
      oSel.SetText(cDiD->sK_ng[3]);
      //����������� ������� ����������� ����� �� 1
      N_strok++;


      if(cDiD->sK_stv[3]!="")
      {
      	//194	���������� ������� �������, �� �������� ����������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_stv[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_stv[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_stv[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_stv[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sDelta_f[3]!="")
      {
      	//195	���, ������ �� ������� ����� ��������� ������� � ��������� � ���� �������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_f[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sDelta_f[1][0])+(unsigned char)(cDiD->sDelta_f[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sDelta_f[1]).Right(1));
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_f[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sDelta_f[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sSESR_ChRP[3]!="")
      {
      	//206	%, ����������� ������ �� ������������ ����������� ������ ��� ��� �� ���������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sSESR_ChRP[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sSESR_ChRP[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(4)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sSESR_ChRP[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sSESR_ChRP[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sK_ng_ChRP[3]!="")
      {
      	//207	%, ����������� ������������ �� ��������� ���� � ��� � ������� �� ���
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_ChRP[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_ChRP[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
         oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_ChRP[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_ChRP[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sfi_tau_int_PRP[3]!="")
      {
      	//306	����������� ���������� � �������� ����������������� ��������� ��� ���
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sfi_tau_int_PRP[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sfi_tau_int_PRP[1][0])+(unsigned char)(cDiD->sfi_tau_int_PRP[1][1])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.InsertSymbol((256*(unsigned char)(cDiD->sfi_tau_int_PRP[1][2])+(unsigned char)(cDiD->sfi_tau_int_PRP[1][3])), COleVariant("Times New Roman"), COleVariant(short(True)), COleVariant(short(wdFontBiasDefault)));
      	oSel.SetText((cDiD->sfi_tau_int_PRP[1]).Right(7));
      	oSel.MoveLeft(COleVariant(short(wdCharacter)),COleVariant(long(2)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sfi_tau_int_PRP[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sfi_tau_int_PRP[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sSESR_PRP[3]!="")
      {
      	//307	%, ����������� ������ �� ������������ ����������� ������ ��� ���
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sSESR_PRP[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sSESR_PRP[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(4)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sSESR_PRP[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sSESR_PRP[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sK_ng_PRP[3]!="")
      {
      	//308	%, ����������� ������������  ��� ���
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_PRP[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_PRP[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_PRP[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_PRP[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }


      if(cDiD->sSESR_norm[3]!="")
      {
      	//309	%, ����� �� ����������� ������ �� ������������ ����������� ������ ��� ��������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sSESR_norm[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sSESR_norm[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(4)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sSESR_norm[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sSESR_norm[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }



      if(cDiD->sK_ng_norm[3]!="")
      {
      	//310	%, ����� �� ����������� ������������ ��� ��������� ������
      	//��������� ������
      	oSel.SelectRow();
      	oSel.InsertRowsBelow(COleVariant(short(1)));
      	//��������� ����� "������������ ���������"
      	oCell = oTable.Cell(N_strok+2,2);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_norm[0]);
      	//��������� ����� "�����������"
      	oCell = oTable.Cell(N_strok+2,3);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_norm[1]);
      	oSel.Collapse(COleVariant(short(wdCollapseStart)));
      	oSel.MoveRight(COleVariant(short(wdCharacter)),COleVariant(long(1)),COleVariant(short(wdMove)));
      	oSel.MoveRight(COleVariant(short(wdSentence)),COleVariant(long(1)),COleVariant(short(wdExtend)));
      	oFont.SetSubscript(True);
      	//��������� ����� "�������."
      	oCell = oTable.Cell(N_strok+2,4);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_norm[2]);
      	//��������� ����� "��������"
      	oCell = oTable.Cell(N_strok+2,5);
      	oCell.Select();
      	oSel.SetText(cDiD->sK_ng_norm[3]);
      	//����������� ������� ����������� ����� �� 1
      	N_strok++;
      }
   }

   //������ ������ ������� ����� ������ ������� ������
   oRow=oRows.Item(1);
   oBorders=oRow.GetBorders();
   oBorder=oBorders.Item(wdBorderBottom);
   oBorder.SetLineWidth(wdLineWidth150pt);

   return TRUE;
}//����� Draw_otchet_DMR(cData_interval_DMR &cDiD)

