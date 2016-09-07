//������� ������ ������ ������� �4 � �������� �������� �� ����� 5


#include <afx.h>
#include "afxdisp.h"//���������� ��� ������ AfxOleInit()
#include "msword9.h"//������������ ����, ���������� � ������� ClassWizard Visual Studio
#include "DMRWord.h"//������������ ���� � ���������� ������� MS Word



cBlank_A4_f5::cBlank_A4_f5()//������ �����������
{
}





//�������������� ������� OLE. ������������ ��������� ��������� ���������� OLE.
//��������� ������ ������������� "Word.Application".
//���������� �������� ������� � ����� � ������ ����������� �� ������� ����� ��������� ��������� Word
//�������� ������ ��� ���������� �������� ������� ����� �� ������ �� ������ ���� cData_interval_DMR
//�������� ������ cData_interval_DMR �������� � � ����� DMR.h
//� ������ ��������� ���������� ������ ���������� TRUE, � ������ ������� ��� ����������
//������������� - FALSE.
BOOL	cBlank_A4_f5::Draw_blank(cData_interval_DMR *cDiD)
{




   //��� �������� ��������������������� ������� OLE. ���� ����� �� �������, �� ����� CreateDispatch �� ���������.
   /*������ ���������� �����������. � DLL ������ ������
   if(!AfxOleInit()) // Your addition starts here
   {
   	AfxMessageBox((LPCTSTR)"Could not initialize COM dll",(UINT)MB_OK,(UINT)0);
      return FALSE;
   }               // End of your addition
   */

   AfxEnableControlContainer();//Call this function in application object's InitInstance function to enable support for
      									//containment of OLE controls.

   if(!app.CreateDispatch("Word.Application")) //��������� ������
   {
   	AfxMessageBox((LPCTSTR)"������ ��� ������ Word�!",(UINT)MB_OK,(UINT)0);
      return FALSE;
   }


   app.SetVisible(TRUE); //� ������� Word �������


   COleVariant  covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
   //���� ��������� ����������
   oDocs = app.GetDocuments();
   //�������� � ��� ����� ��������
   //��������! ���� � ��� Word 97 - �� ������� ����� �����:
   //oDocs.Add(covOptional,covOptional);				//97
   oDocs.Add(covOptional,covOptional,covOptional,covOptional);	//2000
   //� �������� ��� ��� �������� ��������� � ������� 1
   Word_blank = oDocs.Item(COleVariant(long(1)));
   //�������������� ��������
   Word_blank.Activate();

   //��������� ���������� ��������
   CPageSetup oPageSetup;
   oPageSetup=this->Word_blank.GetPageSetup();
   oPageSetup.put_DifferentFirstPageHeaderFooter(long(True));//������ ����������� ������� � ��������� ������
   oPageSetup.put_LeftMargin(20.*MM2PH);//���������� �� ������ ���� �����
   oPageSetup.put_RightMargin(5.*MM2PH);//���������� �� ������� ���� �����
   oPageSetup.put_TopMargin(10.*MM2PV);//���������� �� �������� ���� �����
   oPageSetup.put_BottomMargin(5.*MM2PV);//���������� �� ������� ���� �����
   oPageSetup.put_HeaderDistance(0.*MM2PV);//���������� �� �������� ���� ����� �� �������� �����������
   oPageSetup.put_FooterDistance(5.*MM2PV);//���������� �� ������� ���� ����� �� ������� �����������
   Word_blank.SetPageSetup(oPageSetup);

   //��������� ���� 100%
   CWindow0 ActiveWindow;
   ActiveWindow=this->Word_blank.GetActiveWindow();
   CPane ActivePane;
   ActivePane=ActiveWindow.get_ActivePane();
   CView0 View;
   View=ActivePane.get_View();
   CZoom Zoom;
   Zoom=View.get_Zoom();
   Zoom.put_Percentage(100);

   //���� � ������ ���������� (footer)
   View.put_SeekView(wdSeekCurrentPageFooter);

   //������� ����� �������
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

   //������� ������� �������� ������
   Tables oTables;
   Table oTable;
   oRan = oTextFrame.get_TextRange();
   oTables = this->Word_blank.GetTables();

   //�������� ������� � ���������
   oTable = oTables.Add(oRan,7,5,COleVariant(short(wdWord9TableBehavior)),COleVariant(short(wdAutoFitFixed)));
   //��������� ����������� ������ � �������
   oTable.Select();
   oSel.SetOrientation(wdTextOrientationUpward);

   //��������� ������ � �������
   _Font oFont;
   oFont=oSel.GetFont();
   oFont.SetSize(10);
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

   //������������� ������ ��������
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

   //���������� ������
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

   //����������� ������� � ������
   //����� "���. � ����."
   oCell = oTable.Cell(7,2);
   oRan = oCell.GetRange();
   oRan.SetText("���. � ����.");
   oCell.Select();
   oFont.SetSize(8);
   oSel.BoldRun();
   Paragraphs oPars;
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);

   //����� 20
   oCell = oTable.Cell(7,3);
   oRan = oCell.GetRange();
   oRan.SetText(cDiD->sN_podl);
   oCell.Select();
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);

   //����� "����. � ����"
   oCell = oTable.Cell(6,2);
   oRan = oCell.GetRange();
   oRan.SetText("����. � ����");
   oCell.Select();
   oFont.SetSize(8);
   oSel.BoldRun();
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);

   //����� "����. ���. �"
   oCell = oTable.Cell(5,2);
   oRan = oCell.GetRange();
   oRan.SetText("����. ���. �");
   oCell.Select();
   oFont.SetSize(8);
   oSel.BoldRun();
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);

   //����� 22
   oCell = oTable.Cell(5,3);
   oRan = oCell.GetRange();
   oRan.SetText(cDiD->sN_star_podl);
   oCell.Select();
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);

   //����� "�����������"
   oCell = oTable.Cell(1,1);
   oRan = oCell.GetRange();
   oRan.SetText("�����������");
   oCell.Select();
   oFont.SetSize(8);
   oSel.BoldRun();
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_LEFT);
   if(cDiD->sFam_gip!="")
   {
   	//����� 10 � ������� ������
   	oCell = oTable.Cell(4,2);
   	oRan = oCell.GetRange();
   	oRan.SetText("���");
   	oCell.Select();
   	oPars=oRan.GetParagraphs();
   	oPars.SetAlignment(AL_LEFT);
      //����� 11 � ������� ������
   	oCell = oTable.Cell(3,2);
   	oRan = oCell.GetRange();
   	oRan.SetText(cDiD->sFam_gip);
   	oCell.Select();
   	oPars=oRan.GetParagraphs();
   	oPars.SetAlignment(AL_LEFT);
   }


   //������� ������� ��������� ������
   View.put_SeekView(wdSeekCurrentPageFooter);
   oRan = oSel.GetRange();
   oTables = this->Word_blank.GetTables();

   //�������� ������� � ���������
   oTable = oTables.Add(oRan,8,10,COleVariant(short(wdWord9TableBehavior)),COleVariant(short(wdAutoFitFixed)));
   //��������� ������ � �������
   oTable.Select();
   oFont.SetSize(10);
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
   oTable.Select();
   oCells=oSel.GetCells();
   oCells.SetVerticalAlignment(wdCellAlignVerticalCenter);

   //������������� ������ �����
   oRows=oTable.GetRows();
   oRows.SetHeight(5.*MM2PV*kMM2PV);

   //������������� ������ ��������
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

   //���������� ������
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

   //������������� ������� ������ ����� �������
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

   //����������� ������� � ������
   //����� "���."
   oCell = oTable.Cell(3,1);
   oRan = oCell.GetRange();
   oRan.SetText("���.");
   oCell.Select();
   oFont.SetSize(7);
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);

   //����� "���. ��."
   oCell = oTable.Cell(3,2);
   oRan = oCell.GetRange();
   oRan.SetText("���. ��.");
   oCell.Select();
   oFont.SetSize(7);
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);

   //����� "����" � ������� ���������
   oCell = oTable.Cell(3,3);
   oRan = oCell.GetRange();
   oRan.SetText("����");
   oCell.Select();
   oFont.SetSize(7);
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);

   //����� "� ���."
   oCell = oTable.Cell(3,4);
   oRan = oCell.GetRange();
   oRan.SetText("� ���.");
   oCell.Select();
   oFont.SetSize(7);
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);

   //����� "����." � ������� ���������
   oCell = oTable.Cell(3,5);
   oRan = oCell.GetRange();
   oRan.SetText("����.");
   oCell.Select();
   oFont.SetSize(7);
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);

   //����� "����" � ������� ���������
   oCell = oTable.Cell(3,6);
   oRan = oCell.GetRange();
   oRan.SetText("����");
   oCell.Select();
   oFont.SetSize(7);
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);
   if(cDiD->sN_izm!="")
   {
   	//����� 14
   	oCell = oTable.Cell(2,1);
   	oRan = oCell.GetRange();
   	oRan.SetText(cDiD->sN_izm);
   	oCell.Select();
      oFont.SetSize(8);
   	oPars=oRan.GetParagraphs();
   	oPars.SetAlignment(AL_CENTER);
      //����� 15
   	oCell = oTable.Cell(2,2);
   	oRan = oCell.GetRange();
   	oRan.SetText(cDiD->sKol_uch);
   	oCell.Select();
      oFont.SetSize(8);
   	oPars=oRan.GetParagraphs();
   	oPars.SetAlignment(AL_CENTER);
      //����� 16
   	oCell = oTable.Cell(2,3);
   	oRan = oCell.GetRange();
   	oRan.SetText(cDiD->sZam_nov_vse);
   	oCell.Select();
      oFont.SetSize(8);
   	oPars=oRan.GetParagraphs();
   	oPars.SetAlignment(AL_CENTER);
      //����� 17
   	oCell = oTable.Cell(2,4);
   	oRan = oCell.GetRange();
   	oRan.SetText(cDiD->sOb_razr);
   	oCell.Select();
      oFont.SetSize(8);
   	oPars=oRan.GetParagraphs();
   	oPars.SetAlignment(AL_CENTER);
   }

   //����� "������"
   oCell = oTable.Cell(4,1);
   oRan = oCell.GetRange();
   oRan.SetText("������.");
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_LEFT);

   //����� 11 ������� �����������
   oCell = oTable.Cell(4,2);
   oRan = oCell.GetRange();
   oRan.SetText(cDiD->sFam_isp);
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_LEFT);
   if(cDiD->sFam_prov!="")
   {
   	//����� "��������"
   	oCell = oTable.Cell(5,1);
   	oRan = oCell.GetRange();
   	oRan.SetText("��������");
   	oPars=oRan.GetParagraphs();
   	oPars.SetAlignment(AL_LEFT);
   	//����� 11 ������� ������������
   	oCell = oTable.Cell(5,2);
   	oRan = oCell.GetRange();
   	oRan.SetText(cDiD->sFam_prov);
   	oPars=oRan.GetParagraphs();
   	oPars.SetAlignment(AL_LEFT);
   }

   if(cDiD->sFam_glt!="")
   {
   	//����� "��. ����."
   	oCell = oTable.Cell(6,1);
   	oRan = oCell.GetRange();
   	oRan.SetText("��. ����.");
   	oPars=oRan.GetParagraphs();
   	oPars.SetAlignment(AL_LEFT);
   	//����� 11 ������� �������� ���������
   	oCell = oTable.Cell(6,2);
   	oRan = oCell.GetRange();
   	oRan.SetText(cDiD->sFam_glt);
   	oPars=oRan.GetParagraphs();
   	oPars.SetAlignment(AL_LEFT);
   }

   //����� "�. �����."
   oCell = oTable.Cell(7,1);
   oRan = oCell.GetRange();
   oRan.SetText("�. �����.");
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_LEFT);

   //����� 11 ������� ���������������
   oCell = oTable.Cell(7,2);
   oRan = oCell.GetRange();
   oRan.SetText(cDiD->sFam_nk);
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_LEFT);

   if(cDiD->sFam_no!="")
   {
   	//����� "���. ���."
   	oCell = oTable.Cell(8,1);
   	oRan = oCell.GetRange();
   	oRan.SetText("���. ���.");
   	oPars=oRan.GetParagraphs();
   	oPars.SetAlignment(AL_LEFT);
   	//����� 11 ������� ���������� ������
   	oCell = oTable.Cell(8,2);
   	oRan = oCell.GetRange();
   	oRan.SetText(cDiD->sFam_no);
   	oPars=oRan.GetParagraphs();
   	oPars.SetAlignment(AL_LEFT);
   }

   //����� 1
   oCell = oTable.Cell(1,7);
   oRan = oCell.GetRange();
   oRan.SetText(cDiD->sObozn_doc);
   oCell.Select();
   oFont.SetSize(14);
   oSel.BoldRun();
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);

   //����� 5
   oCell = oTable.Cell(4,5);
   oRan = oCell.GetRange();
   oRan.SetText(cDiD->sNaim_doc);
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);

   //����� "������"
   oCell = oTable.Cell(4,6);
   oRan = oCell.GetRange();
   oRan.SetText("������");
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);

   //����� "����"
   oCell = oTable.Cell(4,7);
   oRan = oCell.GetRange();
   oRan.SetText("����");
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);

   //����� "������"
   oCell = oTable.Cell(4,8);
   oRan = oCell.GetRange();
   oRan.SetText("������");
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);

   //����� 6
   oCell = oTable.Cell(5,6);
   oRan = oCell.GetRange();
   oRan.SetText(cDiD->sVid_doc);
   oCell.Select();
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);

   //����� 7
   CString sN_list_pref="";//������� ��������� �������
   long lsn;//��������� ����� �������

   int index;
   if((index=cDiD->sN_list.Find('-'))!=-1)//���� ������ �������� ������ '-'
   {
   	sN_list_pref=cDiD->sN_list.Left(index+1);//������� ��������� �������
      lsn=atol(&((LPCTSTR(cDiD->sN_list))[index+1]));//��������� ����� �������
   }
   else lsn=atol(LPCTSTR(cDiD->sN_list));//��������� ����� �������
   oCell = oTable.Cell(5,7);
   oCell.Select();
   oRan = oSel.GetRange();
   oSel.Collapse(COleVariant(short(wdCollapseStart)));
   oSel.SetText(sN_list_pref);
   oSel.Collapse(COleVariant(short(wdCollapseEnd)));
   CFields oFields;
   oFields=oSel.GetFields();
   oFields.Add(oSel.GetRange(), COleVariant(short(wdFieldPage)), covOptional, covOptional);//������� ���� "����� ��������"
   CPageNumbers oPN;
   oPN=oHeaderFooter.get_PageNumbers();
   oPN.put_RestartNumberingAtSection(TRUE);
   oPN.put_StartingNumber(lsn);
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);

   //����� 8
   oCell = oTable.Cell(5,8);
   oRan = oCell.GetRange();
   oRan.SetText(cDiD->sKol_list);
   oCell.Select();
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);

   //����� 9
   oCell = oTable.Cell(6,6);
   oRan = oCell.GetRange();
   oRan.SetText(cDiD->sNaim_razr);
   oCell.Select();
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);


   //������ ������������� �����
   oShape=oShapes.AddShape(msoShapeRectangle,20.*2.82,5.*2.787,185.*2.834,287.*2.787,covOptional);
   oFillFormat=oShape.get_Fill();
   oLineFormat=oShape.get_Line();
   oFillFormat.put_Visible(long (FALSE));
   oLineFormat.put_Weight(0.5*MM2PH);


   //��������� ��������
   View.put_SeekView(wdSeekMainDocument);//����� �� �����������
   COleVariant covBreakType((long)BR_PAGE);
   oSel.InsertBreak(covBreakType);


   //���� � ������ ���������� 2-� ��������
   View.put_SeekView(wdSeekCurrentPageFooter);
   oSel.MoveDown(COleVariant(short(wdWindow)),COleVariant(long(1)),COleVariant(short(wdMove)));
   oSel.MoveDown(COleVariant(short(wdLine)),COleVariant(long(2)),COleVariant(short(wdMove)));


   //������� ����� ������� ��� �������� ������ �� 2-� ��������
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


   //������� ������� �������� ������ �� 2-� ��������
   oRan = oTextFrame.get_TextRange();
   oTables = this->Word_blank.GetTables();

   //�������� ������� � ���������
   oTable = oTables.Add(oRan,3,2,COleVariant(short(wdWord9TableBehavior)),COleVariant(short(wdAutoFitFixed)));
   //��������� ����������� ������ � �������
   oTable.Select();
   oSel.SetOrientation(wdTextOrientationUpward);

   //��������� ������ � �������
   oFont.SetSize(10);
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
   oTable.Select();
   oCells=oSel.GetCells();
   oCells.SetVerticalAlignment(wdCellAlignVerticalCenter);

   //������������� ������ �����
   oRows=oTable.GetRows();
   oRow=oRows.Item(1);
   oRow.SetHeight(25.*2.783, wdRowHeightExactly);
   oRow=oRows.Item(2);
   oRow.SetHeight(35.*2.783, wdRowHeightExactly);
   oRow=oRows.Item(3);
   oRow.SetHeight(25.*2.783, wdRowHeightExactly);

   //������������� ������ ��������
   oColumns=oTable.GetColumns();
   oColumn=oColumns.Item(1);
   oColumn.SetPreferredWidth(5.*MM2PH);
   oColumn=oColumns.Item(2);
   oColumn.SetPreferredWidth(7.*MM2PH);

   //������������� ������� ������ ����� �������
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

   //����������� ������� � ������
   //����� "���. � ����."
   oCell = oTable.Cell(3,1);
   oRan = oCell.GetRange();
   oRan.SetText("���. � ����.");
   oCell.Select();
   oFont.SetSize(8);
   oSel.BoldRun();
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);

   //����� 20
   oCell = oTable.Cell(3,2);
   oRan = oCell.GetRange();
   oRan.SetText(cDiD->sN_podl);
   oCell.Select();
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);

   //����� "����. � ����"
   oCell = oTable.Cell(2,1);
   oRan = oCell.GetRange();
   oRan.SetText("����. � ����");
   oCell.Select();
   oFont.SetSize(8);
   oSel.BoldRun();
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);

   //����� "����. ���. �"
   oCell = oTable.Cell(1,1);
   oRan = oCell.GetRange();
   oRan.SetText("����. ���. �");
   oCell.Select();
   oFont.SetSize(8);
   oSel.BoldRun();
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);

   //����� 22
   oCell = oTable.Cell(1,2);
   oRan = oCell.GetRange();
   oRan.SetText(cDiD->sN_star_podl);
   oCell.Select();
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);


   //������� ������� ��������� (���������� �� ����� 6) ������ �� 2-� ��������
   View.put_SeekView(wdSeekCurrentPageHeader);
   oSel.MoveDown(COleVariant(short(wdLine)),COleVariant(long(2)),COleVariant(short(wdMove)));
   oRan = oSel.GetRange();
   oTables = this->Word_blank.GetTables();

   //�������� ������� � ���������
   oTable = oTables.Add(oRan,4,8,COleVariant(short(wdWord9TableBehavior)),COleVariant(short(wdAutoFitFixed)));
   //��������� ������ � �������
   oTable.Select();
   oFont.SetSize(10);
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
   oTable.Select();
   oCells=oSel.GetCells();
   oCells.SetVerticalAlignment(wdCellAlignVerticalCenter);

   //������������� ������ �����
   oRows=oTable.GetRows();
   oRows.SetHeight(5.*2.771, wdRowHeightExactly);
   oRow=oRows.Item(2);
   oRow.SetHeight(2.*2.771, wdRowHeightExactly);
   oRow=oRows.Item(3);
   oRow.SetHeight(3.*2.771, wdRowHeightExactly);

   //������������� ������ ��������
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

   //���������� ������
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

   //������������� ������� ������ ����� �������
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

   //����������� ������� � ������
   //����� "���."
   oCell = oTable.Cell(4,1);
   oRan = oCell.GetRange();
   oRan.SetText("���.");
   oCell.Select();
   oFont.SetSize(7);
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);

   //����� "���. ��."
   oCell = oTable.Cell(4,2);
   oRan = oCell.GetRange();
   oRan.SetText("���. ��.");
   oCell.Select();
   oFont.SetSize(7);
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);

   //����� "����" � ������� ���������
   oCell = oTable.Cell(4,3);
   oRan = oCell.GetRange();
   oRan.SetText("����");
   oCell.Select();
   oFont.SetSize(7);
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);

   //����� "� ���."
   oCell = oTable.Cell(4,4);
   oRan = oCell.GetRange();
   oRan.SetText("� ���.");
   oCell.Select();
   oFont.SetSize(7);
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);

   //����� "����." � ������� ���������
   oCell = oTable.Cell(4,5);
   oRan = oCell.GetRange();
   oRan.SetText("����.");
   oCell.Select();
   oFont.SetSize(7);
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);

   //����� "����" � ������� ���������
   oCell = oTable.Cell(4,6);
   oRan = oCell.GetRange();
   oRan.SetText("����");
   oCell.Select();
   oFont.SetSize(7);
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);

   //����� 1
   oCell = oTable.Cell(1,7);
   oRan = oCell.GetRange();
   oRan.SetText(cDiD->sObozn_doc);
   oCell.Select();
   oFont.SetSize(14);
   oSel.BoldRun();
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);

   //����� "����"
   oCell = oTable.Cell(1,8);
   oRan = oCell.GetRange();
   oRan.SetText("����");
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);

   //����� 7
   oCell = oTable.Cell(3,8);
   oCell.Select();
   oRan = oSel.GetRange();
   oSel.Collapse(COleVariant(short(wdCollapseStart)));
   oSel.SetText(sN_list_pref);
   oSel.Collapse(COleVariant(short(wdCollapseEnd)));
   oFields=oSel.GetFields();
	oFields.Add(oSel.GetRange(), COleVariant(short(wdFieldPage)), covOptional, covOptional);//������� ���� "����� ��������"
   oPars=oRan.GetParagraphs();
   oPars.SetAlignment(AL_CENTER);


   //������ ������������� ����� �� ������ ��������
   oShape=oShapes.AddShape(msoShapeRectangle,20.*2.82,5.*2.787,185.*2.834,287.*2.787,covOptional);
   oFillFormat=oShape.get_Fill();
   oLineFormat=oShape.get_Line();
   oFillFormat.put_Visible(long (FALSE));
   oLineFormat.put_Weight(0.5*MM2PH);


   //������� 2-� ��������
   View.put_SeekView(wdSeekMainDocument);//����� �� �����������
   oSel.MoveUp(COleVariant(short(wdParagraph)),COleVariant(short(1)),COleVariant(short(wdExtend)));
   oSel.Delete(COleVariant(short(wdCharacter)),COleVariant(short(1)));

   return TRUE;
}//BOOL	cBlank_A4_f5::Draw_blank(cData_interval_DMR &cDiD)

