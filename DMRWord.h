#ifndef DMRWORD_H//Чтобы не включался многократно
#define DMRWORD_H

#include	"CWindow0.h"
#include	"CPane.h"
#include	"CView0.h"
#include "CFrames.h"
#include "CFrame.h"
#include "CPageSetup.h"
#include "CTextFrame.h"
#include "CShapeRange.h"
#include "CHeaderFooter.h"
#include "CShapes.h"
#include "CShape.h"
#include "CFillFormat.h"
#include "CLineFormat.h"
#include "CSections.h"
#include "CSection.h"
#include "CParagraphFormat.h"
#include "CTabStops.h"
#include "CZoom.h"
#include "CListGalleries.h"
#include "CListGallery.h"
#include "CListTemplates.h"
#include "CListTemplate.h"
#include "CListLevels.h"
#include "CListLevel.h"
#include "CListFormat.h"
#include "CFields.h"
#include "CPageNumbers.h"

#include "DMR.h"
#include "msword9.h"
//Определение констант MS Word (https://msdn.microsoft.com/en-us/library/office/aa211923(v=office.11).aspx)
//WdOpenFormat
#define	wdOpenFormatAuto			0
#define	wdOpenFormatDocument		1
#define	wdOpenFormatTemplate		2
#define	wdOpenFormatRTF			3
#define	wdOpenFormatText			4
#define	wdOpenFormatUnicodeText	5
#define	wdOpenFormatEncodedText	5
#define	wdOpenFormatAllWord		6
#define	wdOpenFormatWebPages		7
//WdSaveOptions
#define	wdDoNotSaveChanges		0
#define	wdPromptToSaveChanges	-2
#define	wdSaveChanges				-1
//WdOriginalFormat
#define	wdOriginalDocumentFormat	1
#define	wdPromptUser					2
#define	wdWordDocument					0
//WdDefaultTableBehavior
#define	wdWord8TableBehavior	0
#define	wdWord9TableBehavior	1
//WdAutoFitBehavior
#define	wdAutoFitContent	1
#define	wdAutoFitFixed		0
#define	wdAutoFitWindow	2
//WdLineStyle
#define	wdLineStyleNone						0
#define	wdLineStyleSingle						1
#define	wdLineStyleDot							2
#define	wdLineStyleDashSmallGap				3
#define	wdLineStyleDashLargeGap				4
#define	wdLineStyleDashDot					5
#define	wdLineStyleDashDotDot				6
#define	wdLineStyleDouble						7
#define	wdLineStyleTriple						8
#define	wdLineStyleThinThickSmallGap		9
#define	wdLineStyleThickThinSmallGap		10
#define	wdLineStyleThinThickThinSmallGap	11
#define	wdLineStyleThinThickMedGap			12
#define	wdLineStyleThickThinMedGap			13
#define	wdLineStyleThinThickThinMedGap	14
#define	wdLineStyleThinThickLargeGap		15
#define	wdLineStyleThickThinLargeGap		16
#define	wdLineStyleThinThickThinLargeGap	17
#define	wdLineStyleSingleWavy				18
#define	wdLineStyleDoubleWavy				19
#define	wdLineStyleDashDotStroked			20
#define	wdLineStyleEmboss3D					21
#define	wdLineStyleEngrave3D					22
#define	wdLineStyleOutset						23
#define	wdLineStyleInset						24
//WdBorderType
#define	wdBorderBottom			-3
#define	wdBorderDiagonalDown	-7
#define	wdBorderDiagonalUp	-8
#define	wdBorderHorizontal	-5
#define	wdBorderLeft			-2
#define	wdBorderRight			-4
#define	wdBorderTop				-1
#define	wdBorderVertical		-6
//WdUnits
#define	wdCell						12
#define	wdCharacter					1
#define	wdCharacterFormatting	13
#define	wdColumn						9
#define	wdItem						16
#define	wdLine						5
#define	wdParagraph					4
#define	wdParagraphFormatting	14
#define	wdRow							10
#define	wdScreen						7
#define	wdSection					8
#define	wdSentence					3
#define	wdStory						6
#define	wdTable						15
#define	wdWindow						11
#define	wdWord						2
//WdConstants
#define	wdAutoPosition	0
#define	wdBackward		-1073741823
#define	wdCreatorCode	1297307460
#define	wdFirst			1
#define	wdForward		1073741823
#define	wdToggle			9999998
#define	wdUndefined		9999999
//WdMovementType
#define	wdExtend	1
#define	wdMove	0
//WdTableFormat
#define	wdTableFormat3DEffects1		32
#define	wdTableFormat3DEffects2		33
#define	wdTableFormat3DEffects3		34
#define	wdTableFormatClassic1		4
#define	wdTableFormatClassic2		5
#define	wdTableFormatClassic3		6
#define	wdTableFormatClassic4		7
#define	wdTableFormatColorful1		8
#define	wdTableFormatColorful2		9
#define	wdTableFormatColorful3		10
#define	wdTableFormatColumns1		11
#define	wdTableFormatColumns2		12
#define	wdTableFormatColumns3		13
#define	wdTableFormatColumns4		14
#define	wdTableFormatColumns5		15
#define	wdTableFormatContemporary	35
#define	wdTableFormatElegant			36
#define	wdTableFormatGrid1			16
#define	wdTableFormatGrid2			17
#define	wdTableFormatGrid3			18
#define	wdTableFormatGrid4			19
#define	wdTableFormatGrid5			20
#define	wdTableFormatGrid6			21
#define	wdTableFormatGrid7			22
#define	wdTableFormatGrid8			23
#define	wdTableFormatList1			24
#define	wdTableFormatList2			25
#define	wdTableFormatList3			26
#define	wdTableFormatList4			27
#define	wdTableFormatList5			28
#define	wdTableFormatList6			29
#define	wdTableFormatList7			30
#define	wdTableFormatList8			31
#define	wdTableFormatNone				0
#define	wdTableFormatProfessional	37
#define	wdTableFormatSimple1			1
#define	wdTableFormatSimple2			2
#define	wdTableFormatSimple3			3
#define	wdTableFormatSubtle1			38
#define	wdTableFormatSubtle2			39
#define	wdTableFormatWeb1				40
#define	wdTableFormatWeb2				41
#define	wdTableFormatWeb3				42
//WdFindWrap
#define	wdFindAsk		2
#define	wdFindContinue	1
#define	wdFindStop		0
//WdSeekView
#define	wdSeekCurrentPageFooter	10
#define	wdSeekCurrentPageHeader	9
#define	wdSeekEndnotes				8
#define	wdSeekEvenPagesFooter	6
#define	wdSeekEvenPagesHeader	3
#define	wdSeekFirstPageFooter	5
#define	wdSeekFirstPageHeader	2
#define	wdSeekFootnotes			7
#define	wdSeekMainDocument		0
#define	wdSeekPrimaryFooter		4
#define	wdSeekPrimaryHeader		1
//WdFrameSizeRule
#define	wdFrameAtLeast	1
#define	wdFrameAuto		0
#define	wdFrameExact	2
//WdTextOrientation
#define	wdTextOrientationDownward						3
#define	wdTextOrientationHorizontal					0
#define	wdTextOrientationHorizontalRotatedFarEast	4
#define	wdTextOrientationUpward							2
#define	wdTextOrientationVerticalFarEast				1
//WdLineWidth
#define	wdLineWidth025pt	2
#define	wdLineWidth050pt	4
#define	wdLineWidth075pt	6
#define	wdLineWidth100pt	8
#define	wdLineWidth150pt	12
#define	wdLineWidth225pt	18
#define	wdLineWidth300pt	24
#define	wdLineWidth450pt	36
#define	wdLineWidth600pt	48
//WdCellVerticalAlignment
#define	wdCellAlignVerticalBottom	3
#define	wdCellAlignVerticalCenter	1
#define	wdCellAlignVerticalTop		0
//WdRelativeVerticalPosition
#define	wdRelativeVerticalPositionLine		3
#define	wdRelativeVerticalPositionMargin		0
#define	wdRelativeVerticalPositionPage		1
#define	wdRelativeVerticalPositionParagraph	2
//WdRelativeHorizontalPosition
#define	wdRelativeHorizontalPositionCharacter	3
#define	wdRelativeHorizontalPositionColumn		2
#define	wdRelativeHorizontalPositionMargin		0
#define	wdRelativeHorizontalPositionPage		1
//MsoAutoShapeType Enumeration
#define	msoShape16pointStar							94	//16-point star.
#define	msoShape24pointStar							95	//24-point star.
#define	msoShape32pointStar							96	//32-point star.
#define	msoShape4pointStar							91	//4-point star.
#define	msoShape5pointStar							92	//5-point star.
#define	msoShape8pointStar							93	//8-point star.
#define	msoShapeActionButtonBackorPrevious		129	//Back or Previous button. Supports mouse-click and mouse-over actions.
#define	msoShapeActionButtonBeginning				131	//Beginning button. Supports mouse-click and mouse-over actions.
#define	msoShapeActionButtonCustom					125	//Button with no default picture or text. Supports mouse-click and mouse-over actions.
#define	msoShapeActionButtonDocument				134	//Document button. Supports mouse-click and mouse-over actions.
#define	msoShapeActionButtonEnd						132	//End button. Supports mouse-click and mouse-over actions.
#define	msoShapeActionButtonForwardorNext		130	//Forward or Next button. Supports mouse-click and mouse-over actions.
#define	msoShapeActionButtonHelp					127	//Help button. Supports mouse-click and mouse-over actions.
#define	msoShapeActionButtonHome					126	//Home button. Supports mouse-click and mouse-over actions.
#define	msoShapeActionButtonInformation			128	//Information button. Supports mouse-click and mouse-over actions.
#define	msoShapeActionButtonMovie					136	//Movie button. Supports mouse-click and mouse-over actions.
#define	msoShapeActionButtonReturn					133	//Return button. Supports mouse-click and mouse-over actions.
#define	msoShapeActionButtonSound					135	//Sound button. Supports mouse-click and mouse-over actions.
#define	msoShapeArc										25	//Arc.
#define	msoShapeBalloon								137	//Balloon.
#define	msoShapeBentArrow								41	//Block arrow that follows a curved 90-degree angle.
#define	msoShapeBentUpArrow							44	//Block arrow that follows a sharp 90-degree angle. Points up by default.
#define	msoShapeBevel									15	//Bevel.
#define	msoShapeBlockArc								20	//Block arc.
#define	msoShapeCan										13	//Can.
#define	msoShapeChevron								52	//Chevron.
#define	msoShapeCircularArrow						60	//Block arrow that follows a curved 180-degree angle.
#define	msoShapeCloudCallout							108	//Cloud callout.
#define	msoShapeCross									11	//Cross.
#define	msoShapeCube									14	//Cube.
#define	msoShapeCurvedDownArrow						48	//Block arrow that curves down.
#define	msoShapeCurvedDownRibbon					100	//Ribbon banner that curves down.
#define	msoShapeCurvedLeftArrow						46	//Block arrow that curves left.
#define	msoShapeCurvedRightArrow					45	//Block arrow that curves right.
#define	msoShapeCurvedUpArrow						47	//Block arrow that curves up.
#define	msoShapeCurvedUpRibbon						99	//Ribbon banner that curves up.
#define	msoShapeDiamond								4	//Diamond.
#define	msoShapeDonut									18	//Donut.
#define	msoShapeDoubleBrace							27	//Double brace.
#define	msoShapeDoubleBracket						26	//Double bracket.
#define	msoShapeDoubleWave							104	//Double wave.
#define	msoShapeDownArrow								36	//Block arrow that points down.
#define	msoShapeDownArrowCallout					56	//Callout with arrow that points down.
#define	msoShapeDownRibbon							98	//Ribbon banner with center area below ribbon ends.
#define	msoShapeExplosion1							89	//Explosion.
#define	msoShapeExplosion2							90	//Explosion.
#define	msoShapeFlowchartAlternateProcess		62	//Alternate process flowchart symbol.
#define	msoShapeFlowchartCard						75	//Card flowchart symbol.
#define	msoShapeFlowchartCollate					79	//Collate flowchart symbol.
#define	msoShapeFlowchartConnector					73	//Connector flowchart symbol.
#define	msoShapeFlowchartData						64	//Data flowchart symbol.
#define	msoShapeFlowchartDecision					63	//Decision flowchart symbol.
#define	msoShapeFlowchartDelay						84	//Delay flowchart symbol.
#define	msoShapeFlowchartDirectAccessStorage	87	//Direct access storage flowchart symbol.
#define	msoShapeFlowchartDisplay					88	//Display flowchart symbol.
#define	msoShapeFlowchartDocument					67	//Document flowchart symbol.
#define	msoShapeFlowchartExtract					81	//Extract flowchart symbol.
#define	msoShapeFlowchartInternalStorage			66	//Internal storage flowchart symbol.
#define	msoShapeFlowchartMagneticDisk				86	//Magnetic disk flowchart symbol.
#define	msoShapeFlowchartManualInput				71	//Manual input flowchart symbol.
#define	msoShapeFlowchartManualOperation			72	//Manual operation flowchart symbol.
#define	msoShapeFlowchartMerge						82	//Merge flowchart symbol.
#define	msoShapeFlowchartMultidocument			68	//Multi-document flowchart symbol.
#define	msoShapeFlowchartOffpageConnector		74	//Off-page connector flowchart symbol.
#define	msoShapeFlowchartOr							78	//"Or" flowchart symbol.
#define	msoShapeFlowchartPredefinedProcess		65	//Predefined process flowchart symbol.
#define	msoShapeFlowchartPreparation				70	//Preparation flowchart symbol.
#define	msoShapeFlowchartProcess					61	//Process flowchart symbol.
#define	msoShapeFlowchartPunchedTape				76	//Punched tape flowchart symbol.
#define	msoShapeFlowchartSequentialAccessStorage	//85	Sequential access storage flowchart symbol.
#define	msoShapeFlowchartSort						80	//Sort flowchart symbol.
#define	msoShapeFlowchartStoredData				83	//Stored data flowchart symbol.
#define	msoShapeFlowchartSummingJunction			77	//Summing junction flowchart symbol.
#define	msoShapeFlowchartTerminator				69	//Terminator flowchart symbol.
#define	msoShapeFoldedCorner							16	//Folded corner.
#define	msoShapeHeart									21	//Heart.
#define	msoShapeHexagon								10	//Hexagon.
#define	msoShapeHorizontalScroll					102	//Horizontal scroll.
#define	msoShapeIsoscelesTriangle					7	//Isosceles triangle.
#define	msoShapeLeftArrow								34	//Block arrow that points left.
#define	msoShapeLeftArrowCallout					54	//Callout with arrow that points left.
#define	msoShapeLeftBrace								31	//Left brace.
#define	msoShapeLeftBracket							29	//Left bracket.
#define	msoShapeLeftRightArrow						37	//Block arrow with arrowheads that point both left and right.
#define	msoShapeLeftRightArrowCallout				57	//Callout with arrowheads that point both left and right.
#define	msoShapeLeftRightUpArrow					40	//Block arrow with arrowheads that point left, right, and up.
#define	msoShapeLeftUpArrow							43	//Block arrow with arrowheads that point left and up.
#define	msoShapeLightningBolt						22	//Lightning bolt.
#define	msoShapeLineCallout1							109	//Callout with border and horizontal callout line.
#define	msoShapeLineCallout1AccentBar				113	//Callout with horizontal accent bar.
#define	msoShapeLineCallout1BorderandAccentBar	121	//Callout with border and horizontal accent bar.
#define	msoShapeLineCallout1NoBorder				117	//Callout with horizontal line.
#define	msoShapeLineCallout2							110	//Callout with diagonal straight line.
#define	msoShapeLineCallout2AccentBar				114	//Callout with diagonal callout line and accent bar.
#define	msoShapeLineCallout2BorderandAccentBar	122	//Callout with border, diagonal straight line, and accent bar.
#define	msoShapeLineCallout2NoBorder				118	//Callout with no border and diagonal callout line.
#define	msoShapeLineCallout3							111	//Callout with angled line.
#define	msoShapeLineCallout3AccentBar				115	//Callout with angled callout line and accent bar.
#define	msoShapeLineCallout3BorderandAccentBar	123	//Callout with border, angled callout line, and accent bar.
#define	msoShapeLineCallout3NoBorder				119	//Callout with no border and angled callout line.
#define	msoShapeLineCallout4							112	//Callout with callout line segments forming a U-shape.
#define	msoShapeLineCallout4AccentBar				116	//Callout with accent bar and callout line segments forming a U-shape.
#define	msoShapeLineCallout4BorderandAccentBar	124	//Callout with border, accent bar, and callout line segments forming a U-shape.
#define	msoShapeLineCallout4NoBorder				120	//Callout with no border and callout line segments forming a U-shape.
#define	msoShapeMixed									-2	//Return value only; indicates a combination of the other states.
#define	msoShapeMoon									24	//Moon.
#define	msoShapeNoSymbol								19	//"No" symbol.
#define	msoShapeNotchedRightArrow					50	//Notched block arrow that points right.
#define	msoShapeNotPrimitive							138	//Not supported.
#define	msoShapeOctagon								6	//Octagon.
#define	msoShapeOval									9	//Oval.
#define	msoShapeOvalCallout							107	//Oval-shaped callout.
#define	msoShapeParallelogram						2	//Parallelogram.
#define	msoShapePentagon								51	//Pentagon.
#define	msoShapePlaque									28	//Plaque.
#define	msoShapeQuadArrow								39	//Block arrows that point up, down, left, and right.
#define	msoShapeQuadArrowCallout					59	//Callout with arrows that point up, down, left, and right.
#define	msoShapeRectangle								1	//Rectangle.
#define	msoShapeRectangularCallout					105	//Rectangular callout.
#define	msoShapeRegularPentagon						12	//Pentagon.
#define	msoShapeRightArrow							33	//Block arrow that points right.
#define	msoShapeRightArrowCallout					53	//Callout with arrow that points right.
#define	msoShapeRightBrace							32	//Right brace.
#define	msoShapeRightBracket							30	//Right bracket.
#define	msoShapeRightTriangle						8	//Right triangle.
#define	msoShapeRoundedRectangle					5	//Rounded rectangle.
#define	msoShapeRoundedRectangularCallout		106	//Rounded rectangle-shaped callout.
#define	msoShapeSmileyFace							17	//Smiley face.
#define	msoShapeStripedRightArrow					49	//Block arrow that points right with stripes at the tail.
#define	msoShapeSun										23	//Sun.
#define	msoShapeTrapezoid								3	//Trapezoid.
#define	msoShapeUpArrow								35	//Block arrow that points up.
#define	msoShapeUpArrowCallout						55	//Callout with arrow that points up.
#define	msoShapeUpDownArrow							38	//Block arrow that points up and down.
#define	msoShapeUpDownArrowCallout					58	//Callout with arrows that point up and down.
#define	msoShapeUpRibbon								97	//Ribbon banner with center area above ribbon ends.
#define	msoShapeUTurnArrow							42	//Block arrow forming a U shape.
#define	msoShapeVerticalScroll						101	//Vertical scroll.
#define	msoShapeWave									103	//Wave.
//WdRowHeightRule
#define	wdRowHeightAtLeast	1
#define	wdRowHeightAuto		0
#define	wdRowHeightExactly	2
//WdRulerStyle
#define	wdAdjustFirstColumn	2
#define	wdAdjustNone			0
#define	wdAdjustProportional	1
#define	wdAdjustSameWidth		3
//WdFontBias
#define	wdFontBiasDefault		0
#define	wdFontBiasDontCare	255
#define	wdFontBiasFareast		1
//WdCollapseDirection
#define	wdCollapseEnd		0
#define	wdCollapseStart	1
//WdTabAlignment
#define	wdAlignTabBar		4
#define	wdAlignTabCenter	1
#define	wdAlignTabDecimal	3
#define	wdAlignTabLeft		0
#define	wdAlignTabList		6
#define	wdAlignTabRight	2
//WdTabLeader
#define	wdTabLeaderDashes		2
#define	wdTabLeaderDots		1
#define	wdTabLeaderHeavy		4
#define	wdTabLeaderLines		3
#define	wdTabLeaderMiddleDot	5
#define	wdTabLeaderSpaces		0
//WdParagraphAlignment
#define	wdAlignParagraphCenter			1
#define	wdAlignParagraphDistribute		4
#define	wdAlignParagraphJustify			3
#define	wdAlignParagraphJustifyHi		7
#define	wdAlignParagraphJustifyLow		8
#define	wdAlignParagraphJustifyMed		5
#define	wdAlignParagraphLeft				0
#define	wdAlignParagraphRight			2
#define	wdAlignParagraphThaiJustify	9
//WdListGalleryType
#define	wdBulletGallery			1
#define	wdNumberGallery			2
#define	wdOutlineNumberGallery	3
//WdListApplyTo
#define	wdListApplyToSelection			2
#define	wdListApplyToThisPointForward	1
#define	wdListApplyToWholeList			0
//WdDefaultListBehavior
#define	wdWord10ListBehavior	2
#define	wdWord8ListBehavior	0
#define	wdWord9ListBehavior	1
//WdFieldType
#define	wdFieldAddin	81
#define	wdFieldAddressBlock	93
#define	wdFieldAdvance	84
#define	wdFieldAsk	38
#define	wdFieldAuthor	17
#define	wdFieldAutoNum	54
#define	wdFieldAutoNumLegal	53
#define	wdFieldAutoNumOutline	52
#define	wdFieldAutoText	79
#define	wdFieldAutoTextList	89
#define	wdFieldBarCode	63
#define	wdFieldBidiOutline	92
#define	wdFieldComments	19
#define	wdFieldCompare	80
#define	wdFieldCreateDate	21
#define	wdFieldData	40
#define	wdFieldDatabase	78
#define	wdFieldDate	31
#define	wdFieldDDE	45
#define	wdFieldDDEAuto	46
#define	wdFieldDocProperty	85
#define	wdFieldDocVariable	64
#define	wdFieldEditTime	25
#define	wdFieldEmbed	58
#define	wdFieldEmpty	-1
#define	wdFieldExpression	34
#define	wdFieldFileName	29
#define	wdFieldFileSize	69
#define	wdFieldFillIn	39
#define	wdFieldFootnoteRef	5
#define	wdFieldFormCheckBox	71
#define	wdFieldFormDropDown	83
#define	wdFieldFormTextInput	70
#define	wdFieldFormula	49
#define	wdFieldGlossary	47
#define	wdFieldGoToButton	50
#define	wdFieldGreetingLine	94
#define	wdFieldHTMLActiveX	91
#define	wdFieldHyperlink	88
#define	wdFieldIf	7
#define	wdFieldImport	55
#define	wdFieldInclude	36
#define	wdFieldIncludePicture	67
#define	wdFieldIncludeText	68
#define	wdFieldIndex	8
#define	wdFieldIndexEntry	4
#define	wdFieldInfo	14
#define	wdFieldKeyWord	18
#define	wdFieldLastSavedBy	20
#define	wdFieldLink	56
#define	wdFieldListNum	90
#define	wdFieldMacroButton	51
#define	wdFieldMergeField	59
#define	wdFieldMergeRec	44
#define	wdFieldMergeSeq	75
#define	wdFieldNext	41
#define	wdFieldNextIf	42
#define	wdFieldNoteRef	72
#define	wdFieldNumChars	28
#define	wdFieldNumPages	26
#define	wdFieldNumWords	27
#define	wdFieldOCX	87
#define	wdFieldPage	33
#define	wdFieldPageRef	37
#define	wdFieldPrint	48
#define	wdFieldPrintDate	23
#define	wdFieldPrivate	77
#define	wdFieldQuote	35
#define	wdFieldRef	3
#define	wdFieldRefDoc	11
#define	wdFieldRevisionNum	24
#define	wdFieldSaveDate	22
#define	wdFieldSection	65
#define	wdFieldSectionPages	66
#define	wdFieldSequence	12
#define	wdFieldSet	6
#define	wdFieldShape	95
#define	wdFieldSkipIf	43
#define	wdFieldStyleRef	10
#define	wdFieldSubject	16
#define	wdFieldSubscriber	82
#define	wdFieldSymbol	57
#define	wdFieldTemplate	30
#define	wdFieldTime	32
#define	wdFieldTitle	15
#define	wdFieldTOA	73
#define	wdFieldTOAEntry	74
#define	wdFieldTOC	13
#define	wdFieldTOCEntry	9
#define	wdFieldUserAddress	62
#define	wdFieldUserInitials	61
#define	wdFieldUserName	60



//Константы, необходимые для задания типов выравнивания и вставки разрыва страницы и секции
#define AL_LEFT	0
#define AL_CENTER	1
#define AL_RIGHT	2
#define AL_JUST	3

#define BR_PAGE	0
#define BR_SECT	1

#define MM2PH	2.85	//Миллиметры переводим в горизонтальные пойнты (если таблица рисуется в Frame)
#define MM2PV	2.714	//Миллиметры переводим в вертикальные пойнты (если таблица рисуется в Frame)
#define kMM2PV	0.919 //поправка к MM2PV (если таблица рисуется не в Frame)

//Для функций библиотеки типов MS Word
#define True	-1
#define False	0

//Вид отчета
#define KRATKIY	true
#define POLNY		false


//Класс бланка формата А4 с основной надписью по форме 5
class cBlank_A4_f5
{
   public:
   	_Application app;//Объект приложения MS Word
   	/*В Word'е все документы являются членами коллекции Documents. Прежде, чем начинать работу с документом
   	(или вообще с элементом коллекции), надо коллекцию получить, элемент добавить, а затем получить добавленный элемент.*/
   	Documents oDocs;
   	_Document Word_blank;//Отображенный с помощью Draw_blank документ

   	cBlank_A4_f5();//пустой конструктор
   	BOOL	Draw_blank(cData_interval_DMR *cDiD);//Отображает основную надпись и рамку в колонтитуле на текущем листе открытого
   														 	 //документа Word. Параметры берет по ссылке на объект типа cData_interval_DMR,
                                                 //описанный в "DMR.h"
                                                 //В случае успеха возвращает TRUE, иначе - FALSE.
};

//Класс отчета программы DMR
class otchet_DMR : public cBlank_A4_f5
{
   public:
      otchet_DMR();//конструктор
   	BOOL	Draw_otchet_DMR(cData_interval_DMR *cDiD);//Отображает отчет результатов расчета качественных показателей ЦРРЛ
   														 	 //в краткой или полной форме. Параметры берет по ссылке на объект
                                                 //типа cData_interval_DMR, описанный в "DMR.h"
                                                 //В случае успеха возвращает TRUE, иначе - FALSE.
};

#endif //#ifndef DMRWORD_H
