// Machine generated IDispatch wrapper class(es) created with ClassWizard
/////////////////////////////////////////////////////////////////////////////
// _Application wrapper class

#ifndef msword9_h//����� �� ��������� �����������
#define msword9_h

class _Application : public COleDispatchDriver
{
public:
	_Application() {}		// Calls COleDispatchDriver default constructor
	_Application(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	_Application(const _Application& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	CString GetName();
	LPDISPATCH GetDocuments();
	LPDISPATCH GetWindows();
	LPDISPATCH GetActiveDocument();
	LPDISPATCH GetActiveWindow();
	LPDISPATCH GetSelection();
	LPDISPATCH GetWordBasic();
	LPDISPATCH GetRecentFiles();
	LPDISPATCH GetNormalTemplate();
	LPDISPATCH GetSystem();
	LPDISPATCH GetAutoCorrect();
	LPDISPATCH GetFontNames();
	LPDISPATCH GetLandscapeFontNames();
	LPDISPATCH GetPortraitFontNames();
	LPDISPATCH GetLanguages();
	LPDISPATCH GetAssistant();
	LPDISPATCH GetBrowser();
	LPDISPATCH GetFileConverters();
	LPDISPATCH GetMailingLabel();
	LPDISPATCH GetDialogs();
	LPDISPATCH GetCaptionLabels();
	LPDISPATCH GetAutoCaptions();
	LPDISPATCH GetAddIns();
	BOOL GetVisible();
	void SetVisible(BOOL bNewValue);
	CString GetVersion();
	BOOL GetScreenUpdating();
	void SetScreenUpdating(BOOL bNewValue);
	BOOL GetPrintPreview();
	void SetPrintPreview(BOOL bNewValue);
	LPDISPATCH GetTasks();
	BOOL GetDisplayStatusBar();
	void SetDisplayStatusBar(BOOL bNewValue);
	BOOL GetSpecialMode();
	long GetUsableWidth();
	long GetUsableHeight();
	BOOL GetMathCoprocessorAvailable();
	BOOL GetMouseAvailable();
	VARIANT GetInternational(long Index);
	CString GetBuild();
	BOOL GetCapsLock();
	BOOL GetNumLock();
	CString GetUserName_();
	void SetUserName(LPCTSTR lpszNewValue);
	CString GetUserInitials();
	void SetUserInitials(LPCTSTR lpszNewValue);
	CString GetUserAddress();
	void SetUserAddress(LPCTSTR lpszNewValue);
	LPDISPATCH GetMacroContainer();
	BOOL GetDisplayRecentFiles();
	void SetDisplayRecentFiles(BOOL bNewValue);
	LPDISPATCH GetCommandBars();
	LPDISPATCH GetSynonymInfo(LPCTSTR Word, VARIANT* LanguageID);
	LPDISPATCH GetVbe();
	CString GetDefaultSaveFormat();
	void SetDefaultSaveFormat(LPCTSTR lpszNewValue);
	LPDISPATCH GetListGalleries();
	CString GetActivePrinter();
	void SetActivePrinter(LPCTSTR lpszNewValue);
	LPDISPATCH GetTemplates();
	LPDISPATCH GetCustomizationContext();
	void SetCustomizationContext(LPDISPATCH newValue);
	LPDISPATCH GetKeyBindings();
	LPDISPATCH GetKeysBoundTo(long KeyCategory, LPCTSTR Command, VARIANT* CommandParameter);
	LPDISPATCH GetFindKey(long KeyCode, VARIANT* KeyCode2);
	CString GetCaption();
	void SetCaption(LPCTSTR lpszNewValue);
	CString GetPath();
	BOOL GetDisplayScrollBars();
	void SetDisplayScrollBars(BOOL bNewValue);
	CString GetStartupPath();
	void SetStartupPath(LPCTSTR lpszNewValue);
	long GetBackgroundSavingStatus();
	long GetBackgroundPrintingStatus();
	long GetLeft();
	void SetLeft(long nNewValue);
	long GetTop();
	void SetTop(long nNewValue);
	long GetWidth();
	void SetWidth(long nNewValue);
	long GetHeight();
	void SetHeight(long nNewValue);
	long GetWindowState();
	void SetWindowState(long nNewValue);
	BOOL GetDisplayAutoCompleteTips();
	void SetDisplayAutoCompleteTips(BOOL bNewValue);
	LPDISPATCH GetOptions();
	long GetDisplayAlerts();
	void SetDisplayAlerts(long nNewValue);
	LPDISPATCH GetCustomDictionaries();
	CString GetPathSeparator();
	void SetStatusBar(LPCTSTR lpszNewValue);
	BOOL GetMAPIAvailable();
	BOOL GetDisplayScreenTips();
	void SetDisplayScreenTips(BOOL bNewValue);
	long GetEnableCancelKey();
	void SetEnableCancelKey(long nNewValue);
	BOOL GetUserControl();
	LPDISPATCH GetFileSearch();
	long GetMailSystem();
	CString GetDefaultTableSeparator();
	void SetDefaultTableSeparator(LPCTSTR lpszNewValue);
	BOOL GetShowVisualBasicEditor();
	void SetShowVisualBasicEditor(BOOL bNewValue);
	CString GetBrowseExtraFileTypes();
	void SetBrowseExtraFileTypes(LPCTSTR lpszNewValue);
	BOOL GetIsObjectValid(LPDISPATCH Object);
	LPDISPATCH GetHangulHanjaDictionaries();
	LPDISPATCH GetMailMessage();
	BOOL GetFocusInMailHeader();
	void Quit(VARIANT* SaveChanges, VARIANT* OriginalFormat, VARIANT* RouteDocument);
	void ScreenRefresh();
	void LookupNameProperties(LPCTSTR Name);
	void SubstituteFont(LPCTSTR UnavailableFont, LPCTSTR SubstituteFont);
	BOOL Repeat(VARIANT* Times);
	void DDEExecute(long Channel, LPCTSTR Command);
	long DDEInitiate(LPCTSTR App, LPCTSTR Topic);
	void DDEPoke(long Channel, LPCTSTR Item, LPCTSTR Data);
	CString DDERequest(long Channel, LPCTSTR Item);
	void DDETerminate(long Channel);
	void DDETerminateAll();
	long BuildKeyCode(long Arg1, VARIANT* Arg2, VARIANT* Arg3, VARIANT* Arg4);
	CString KeyString(long KeyCode, VARIANT* KeyCode2);
	void OrganizerCopy(LPCTSTR Source, LPCTSTR Destination, LPCTSTR Name, long Object);
	void OrganizerDelete(LPCTSTR Source, LPCTSTR Name, long Object);
	void OrganizerRename(LPCTSTR Source, LPCTSTR Name, LPCTSTR NewName, long Object);
	// method 'AddAddress' not emitted because of invalid return type or parameter type
	CString GetAddress(VARIANT* Name, VARIANT* AddressProperties, VARIANT* UseAutoText, VARIANT* DisplaySelectDialog, VARIANT* SelectDialog, VARIANT* CheckNamesDialog, VARIANT* RecentAddressesChoice, VARIANT* UpdateRecentAddresses);
	BOOL CheckGrammar(LPCTSTR String);
	BOOL CheckSpelling(LPCTSTR Word, VARIANT* CustomDictionary, VARIANT* IgnoreUppercase, VARIANT* MainDictionary, VARIANT* CustomDictionary2, VARIANT* CustomDictionary3, VARIANT* CustomDictionary4, VARIANT* CustomDictionary5, 
		VARIANT* CustomDictionary6, VARIANT* CustomDictionary7, VARIANT* CustomDictionary8, VARIANT* CustomDictionary9, VARIANT* CustomDictionary10);
	void ResetIgnoreAll();
	LPDISPATCH GetSpellingSuggestions(LPCTSTR Word, VARIANT* CustomDictionary, VARIANT* IgnoreUppercase, VARIANT* MainDictionary, VARIANT* SuggestionMode, VARIANT* CustomDictionary2, VARIANT* CustomDictionary3, VARIANT* CustomDictionary4, 
		VARIANT* CustomDictionary5, VARIANT* CustomDictionary6, VARIANT* CustomDictionary7, VARIANT* CustomDictionary8, VARIANT* CustomDictionary9, VARIANT* CustomDictionary10);
	void GoBack();
	void Help(VARIANT* HelpType);
	void AutomaticChange();
	void ShowMe();
	void HelpTool();
	LPDISPATCH NewWindow();
	void ListCommands(BOOL ListAllCommands);
	void ShowClipboard();
	void OnTime(VARIANT* When, LPCTSTR Name, VARIANT* Tolerance);
	void NextLetter();
	short MountVolume(LPCTSTR Zone, LPCTSTR Server, LPCTSTR Volume, VARIANT* User, VARIANT* UserPassword, VARIANT* VolumePassword);
	CString CleanString(LPCTSTR String);
	void SendFax();
	void ChangeFileOpenDirectory(LPCTSTR Path);
	void GoForward();
	void Move(long Left, long Top);
	void Resize(long Width, long Height);
	float InchesToPoints(float Inches);
	float CentimetersToPoints(float Centimeters);
	float MillimetersToPoints(float Millimeters);
	float PicasToPoints(float Picas);
	float LinesToPoints(float Lines);
	float PointsToInches(float Points);
	float PointsToCentimeters(float Points);
	float PointsToMillimeters(float Points);
	float PointsToPicas(float Points);
	float PointsToLines(float Points);
	void Activate();
	float PointsToPixels(float Points, VARIANT* fVertical);
	float PixelsToPoints(float Pixels, VARIANT* fVertical);
	void KeyboardLatin();
	void KeyboardBidi();
	void ToggleKeyboard();
	long Keyboard(long LangId);
	CString ProductCode();
	LPDISPATCH DefaultWebOptions();
	void SetDefaultTheme(LPCTSTR Name, long DocumentType);
	CString GetDefaultTheme(long DocumentType);
	LPDISPATCH GetEmailOptions();
	long GetLanguage();
	LPDISPATCH GetCOMAddIns();
	BOOL GetCheckLanguage();
	void SetCheckLanguage(BOOL bNewValue);
	LPDISPATCH GetLanguageSettings();
	LPDISPATCH GetAnswerWizard();
	long GetFeatureInstall();
	void SetFeatureInstall(long nNewValue);
	void PrintOut(VARIANT* Background, VARIANT* Append, VARIANT* Range, VARIANT* OutputFileName, VARIANT* From, VARIANT* To, VARIANT* Item, VARIANT* Copies, VARIANT* Pages, VARIANT* PageType, VARIANT* PrintToFile, VARIANT* Collate, 
		VARIANT* FileName, VARIANT* ActivePrinterMacGX, VARIANT* ManualDuplexPrint, VARIANT* PrintZoomColumn, VARIANT* PrintZoomRow, VARIANT* PrintZoomPaperWidth, VARIANT* PrintZoomPaperHeight);
	VARIANT Run(LPCTSTR MacroName, VARIANT* varg1, VARIANT* varg2, VARIANT* varg3, VARIANT* varg4, VARIANT* varg5, VARIANT* varg6, VARIANT* varg7, VARIANT* varg8, VARIANT* varg9, VARIANT* varg10, VARIANT* varg11, VARIANT* varg12, VARIANT* varg13, 
		VARIANT* varg14, VARIANT* varg15, VARIANT* varg16, VARIANT* varg17, VARIANT* varg18, VARIANT* varg19, VARIANT* varg20, VARIANT* varg21, VARIANT* varg22, VARIANT* varg23, VARIANT* varg24, VARIANT* varg25, VARIANT* varg26, VARIANT* varg27, 
		VARIANT* varg28, VARIANT* varg29, VARIANT* varg30);
};
/////////////////////////////////////////////////////////////////////////////
// Documents wrapper class

class Documents : public COleDispatchDriver
{
public:
	Documents() {}		// Calls COleDispatchDriver default constructor
	Documents(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Documents(const Documents& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	LPUNKNOWN Get_NewEnum();
	long GetCount();
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	LPDISPATCH Item(VARIANT* Index);
	void Close(VARIANT* SaveChanges, VARIANT* OriginalFormat, VARIANT* RouteDocument);
	void Save(VARIANT* NoPrompt, VARIANT* OriginalFormat);
	LPDISPATCH Add(VARIANT* Template, VARIANT* NewTemplate, VARIANT* DocumentType, VARIANT* Visible);
	LPDISPATCH Open(VARIANT* FileName, VARIANT* ConfirmConversions, VARIANT* ReadOnly, VARIANT* AddToRecentFiles, VARIANT* PasswordDocument, VARIANT* PasswordTemplate, VARIANT* Revert, VARIANT* WritePasswordDocument, 
		VARIANT* WritePasswordTemplate, VARIANT* Format, VARIANT* Encoding, VARIANT* Visible);
};
/////////////////////////////////////////////////////////////////////////////
// _Document wrapper class

class _Document : public COleDispatchDriver
{
public:
	_Document() {}		// Calls COleDispatchDriver default constructor
	_Document(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	_Document(const _Document& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	CString GetName();
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	LPDISPATCH GetBuiltInDocumentProperties();
	LPDISPATCH GetCustomDocumentProperties();
	CString GetPath();
	LPDISPATCH GetBookmarks();
	LPDISPATCH GetTables();
	LPDISPATCH GetFootnotes();
	LPDISPATCH GetEndnotes();
	LPDISPATCH GetComments();
	long GetType();
	BOOL GetAutoHyphenation();
	void SetAutoHyphenation(BOOL bNewValue);
	BOOL GetHyphenateCaps();
	void SetHyphenateCaps(BOOL bNewValue);
	long GetHyphenationZone();
	void SetHyphenationZone(long nNewValue);
	long GetConsecutiveHyphensLimit();
	void SetConsecutiveHyphensLimit(long nNewValue);
	LPDISPATCH GetSections();
	LPDISPATCH GetParagraphs();
	LPDISPATCH GetWords();
	LPDISPATCH GetSentences();
	LPDISPATCH GetCharacters();
	LPDISPATCH GetFields();
	LPDISPATCH GetFormFields();
	LPDISPATCH GetStyles();
	LPDISPATCH GetFrames();
	LPDISPATCH GetTablesOfFigures();
	LPDISPATCH GetVariables();
	LPDISPATCH GetMailMerge();
	LPDISPATCH GetEnvelope();
	CString GetFullName();
	LPDISPATCH GetRevisions();
	LPDISPATCH GetTablesOfContents();
	LPDISPATCH GetTablesOfAuthorities();
	LPDISPATCH GetPageSetup();
	void SetPageSetup(LPDISPATCH newValue);
	LPDISPATCH GetWindows();
	BOOL GetHasRoutingSlip();
	void SetHasRoutingSlip(BOOL bNewValue);
	LPDISPATCH GetRoutingSlip();
	BOOL GetRouted();
	LPDISPATCH GetTablesOfAuthoritiesCategories();
	LPDISPATCH GetIndexes();
	BOOL GetSaved();
	void SetSaved(BOOL bNewValue);
	LPDISPATCH GetContent();
	LPDISPATCH GetActiveWindow();
	long GetKind();
	void SetKind(long nNewValue);
	BOOL GetReadOnly();
	LPDISPATCH GetSubdocuments();
	BOOL GetIsMasterDocument();
	float GetDefaultTabStop();
	void SetDefaultTabStop(float newValue);
	BOOL GetEmbedTrueTypeFonts();
	void SetEmbedTrueTypeFonts(BOOL bNewValue);
	BOOL GetSaveFormsData();
	void SetSaveFormsData(BOOL bNewValue);
	BOOL GetReadOnlyRecommended();
	void SetReadOnlyRecommended(BOOL bNewValue);
	BOOL GetSaveSubsetFonts();
	void SetSaveSubsetFonts(BOOL bNewValue);
	BOOL GetCompatibility(long Type);
	void SetCompatibility(long Type, BOOL bNewValue);
	LPDISPATCH GetStoryRanges();
	LPDISPATCH GetCommandBars();
	BOOL GetIsSubdocument();
	long GetSaveFormat();
	long GetProtectionType();
	LPDISPATCH GetHyperlinks();
	LPDISPATCH GetShapes();
	LPDISPATCH GetListTemplates();
	LPDISPATCH GetLists();
	BOOL GetUpdateStylesOnOpen();
	void SetUpdateStylesOnOpen(BOOL bNewValue);
	VARIANT GetAttachedTemplate();
	void SetAttachedTemplate(VARIANT* newValue);
	LPDISPATCH GetInlineShapes();
	LPDISPATCH GetBackground();
	void SetBackground(LPDISPATCH newValue);
	BOOL GetGrammarChecked();
	void SetGrammarChecked(BOOL bNewValue);
	BOOL GetSpellingChecked();
	void SetSpellingChecked(BOOL bNewValue);
	BOOL GetShowGrammaticalErrors();
	void SetShowGrammaticalErrors(BOOL bNewValue);
	BOOL GetShowSpellingErrors();
	void SetShowSpellingErrors(BOOL bNewValue);
	LPDISPATCH GetVersions();
	BOOL GetShowSummary();
	void SetShowSummary(BOOL bNewValue);
	long GetSummaryViewMode();
	void SetSummaryViewMode(long nNewValue);
	long GetSummaryLength();
	void SetSummaryLength(long nNewValue);
	BOOL GetPrintFractionalWidths();
	void SetPrintFractionalWidths(BOOL bNewValue);
	BOOL GetPrintPostScriptOverText();
	void SetPrintPostScriptOverText(BOOL bNewValue);
	LPDISPATCH GetContainer();
	BOOL GetPrintFormsData();
	void SetPrintFormsData(BOOL bNewValue);
	LPDISPATCH GetListParagraphs();
	void SetPassword(LPCTSTR lpszNewValue);
	void SetWritePassword(LPCTSTR lpszNewValue);
	BOOL GetHasPassword();
	BOOL GetWriteReserved();
	CString GetActiveWritingStyle(VARIANT* LanguageID);
	void SetActiveWritingStyle(VARIANT* LanguageID, LPCTSTR lpszNewValue);
	BOOL GetUserControl();
	void SetUserControl(BOOL bNewValue);
	BOOL GetHasMailer();
	void SetHasMailer(BOOL bNewValue);
	LPDISPATCH GetMailer();
	LPDISPATCH GetReadabilityStatistics();
	LPDISPATCH GetGrammaticalErrors();
	LPDISPATCH GetSpellingErrors();
	LPDISPATCH GetVBProject();
	BOOL GetFormsDesign();
	CString Get_CodeName();
	void Set_CodeName(LPCTSTR lpszNewValue);
	CString GetCodeName();
	BOOL GetSnapToGrid();
	void SetSnapToGrid(BOOL bNewValue);
	BOOL GetSnapToShapes();
	void SetSnapToShapes(BOOL bNewValue);
	float GetGridDistanceHorizontal();
	void SetGridDistanceHorizontal(float newValue);
	float GetGridDistanceVertical();
	void SetGridDistanceVertical(float newValue);
	float GetGridOriginHorizontal();
	void SetGridOriginHorizontal(float newValue);
	float GetGridOriginVertical();
	void SetGridOriginVertical(float newValue);
	long GetGridSpaceBetweenHorizontalLines();
	void SetGridSpaceBetweenHorizontalLines(long nNewValue);
	long GetGridSpaceBetweenVerticalLines();
	void SetGridSpaceBetweenVerticalLines(long nNewValue);
	BOOL GetGridOriginFromMargin();
	void SetGridOriginFromMargin(BOOL bNewValue);
	BOOL GetKerningByAlgorithm();
	void SetKerningByAlgorithm(BOOL bNewValue);
	long GetJustificationMode();
	void SetJustificationMode(long nNewValue);
	long GetFarEastLineBreakLevel();
	void SetFarEastLineBreakLevel(long nNewValue);
	CString GetNoLineBreakBefore();
	void SetNoLineBreakBefore(LPCTSTR lpszNewValue);
	CString GetNoLineBreakAfter();
	void SetNoLineBreakAfter(LPCTSTR lpszNewValue);
	BOOL GetTrackRevisions();
	void SetTrackRevisions(BOOL bNewValue);
	BOOL GetPrintRevisions();
	void SetPrintRevisions(BOOL bNewValue);
	BOOL GetShowRevisions();
	void SetShowRevisions(BOOL bNewValue);
	void Close(VARIANT* SaveChanges, VARIANT* OriginalFormat, VARIANT* RouteDocument);
	void SaveAs(VARIANT* FileName, VARIANT* FileFormat, VARIANT* LockComments, VARIANT* Password, VARIANT* AddToRecentFiles, VARIANT* WritePassword, VARIANT* ReadOnlyRecommended, VARIANT* EmbedTrueTypeFonts, VARIANT* SaveNativePictureFormat, 
		VARIANT* SaveFormsData, VARIANT* SaveAsAOCELetter);
	void Repaginate();
	void FitToPages();
	void ManualHyphenation();
	void Select();
	void DataForm();
	void Route();
	void Save();
	void SendMail();
	LPDISPATCH Range(VARIANT* Start, VARIANT* End);
	void RunAutoMacro(long Which);
	void Activate();
	void PrintPreview();
	LPDISPATCH GoTo(VARIANT* What, VARIANT* Which, VARIANT* Count, VARIANT* Name);
	BOOL Undo(VARIANT* Times);
	BOOL Redo(VARIANT* Times);
	long ComputeStatistics(long Statistic, VARIANT* IncludeFootnotesAndEndnotes);
	void MakeCompatibilityDefault();
	void Protect(long Type, VARIANT* NoReset, VARIANT* Password);
	void Unprotect(VARIANT* Password);
	void EditionOptions(long Type, long Option, LPCTSTR Name, VARIANT* Format);
	void RunLetterWizard(VARIANT* LetterContent, VARIANT* WizardMode);
	LPDISPATCH GetLetterContent();
	void SetLetterContent(VARIANT* LetterContent);
	void CopyStylesFromTemplate(LPCTSTR Template);
	void UpdateStyles();
	void CheckGrammar();
	void CheckSpelling(VARIANT* CustomDictionary, VARIANT* IgnoreUppercase, VARIANT* AlwaysSuggest, VARIANT* CustomDictionary2, VARIANT* CustomDictionary3, VARIANT* CustomDictionary4, VARIANT* CustomDictionary5, VARIANT* CustomDictionary6, 
		VARIANT* CustomDictionary7, VARIANT* CustomDictionary8, VARIANT* CustomDictionary9, VARIANT* CustomDictionary10);
	void FollowHyperlink(VARIANT* Address, VARIANT* SubAddress, VARIANT* NewWindow, VARIANT* AddHistory, VARIANT* ExtraInfo, VARIANT* Method, VARIANT* HeaderInfo);
	void AddToFavorites();
	void Reload();
	LPDISPATCH AutoSummarize(VARIANT* Length, VARIANT* Mode, VARIANT* UpdateProperties);
	void RemoveNumbers(VARIANT* NumberType);
	void ConvertNumbersToText(VARIANT* NumberType);
	long CountNumberedItems(VARIANT* NumberType, VARIANT* Level);
	void Post();
	void ToggleFormsDesign();
	void Compare(LPCTSTR Name);
	void UpdateSummaryProperties();
	VARIANT GetCrossReferenceItems(VARIANT* ReferenceType);
	void AutoFormat();
	void ViewCode();
	void ViewPropertyBrowser();
	void ForwardMailer();
	void Reply();
	void ReplyAll();
	void SendMailer(VARIANT* FileFormat, VARIANT* Priority);
	void UndoClear();
	void PresentIt();
	void SendFax(LPCTSTR Address, VARIANT* Subject);
	void Merge(LPCTSTR FileName);
	void ClosePrintPreview();
	void CheckConsistency();
	LPDISPATCH CreateLetterContent(LPCTSTR DateFormat, BOOL IncludeHeaderFooter, LPCTSTR PageDesign, long LetterStyle, BOOL Letterhead, long LetterheadLocation, float LetterheadSize, LPCTSTR RecipientName, LPCTSTR RecipientAddress, 
		LPCTSTR Salutation, long SalutationType, LPCTSTR RecipientReference, LPCTSTR MailingInstructions, LPCTSTR AttentionLine, LPCTSTR Subject, LPCTSTR CCList, LPCTSTR ReturnAddress, LPCTSTR SenderName, LPCTSTR Closing, LPCTSTR SenderCompany, 
		LPCTSTR SenderJobTitle, LPCTSTR SenderInitials, long EnclosureNumber, VARIANT* InfoBlock, VARIANT* RecipientCode, VARIANT* RecipientGender, VARIANT* ReturnAddressShortForm, VARIANT* SenderCity, VARIANT* SenderCode, VARIANT* SenderGender, 
		VARIANT* SenderReference);
	void AcceptAllRevisions();
	void RejectAllRevisions();
	void DetectLanguage();
	void ApplyTheme(LPCTSTR Name);
	void RemoveTheme();
	void WebPagePreview();
	void ReloadAs(long Encoding);
	CString GetActiveTheme();
	CString GetActiveThemeDisplayName();
	LPDISPATCH GetEmail();
	LPDISPATCH GetScripts();
	BOOL GetLanguageDetected();
	void SetLanguageDetected(BOOL bNewValue);
	long GetFarEastLineBreakLanguage();
	void SetFarEastLineBreakLanguage(long nNewValue);
	LPDISPATCH GetFrameset();
	VARIANT GetClickAndTypeParagraphStyle();
	void SetClickAndTypeParagraphStyle(VARIANT* newValue);
	LPDISPATCH GetHTMLProject();
	LPDISPATCH GetWebOptions();
	long GetOpenEncoding();
	long GetSaveEncoding();
	void SetSaveEncoding(long nNewValue);
	BOOL GetOptimizeForWord97();
	void SetOptimizeForWord97(BOOL bNewValue);
	BOOL GetVBASigned();
	void PrintOut(VARIANT* Background, VARIANT* Append, VARIANT* Range, VARIANT* OutputFileName, VARIANT* From, VARIANT* To, VARIANT* Item, VARIANT* Copies, VARIANT* Pages, VARIANT* PageType, VARIANT* PrintToFile, VARIANT* Collate, 
		VARIANT* ActivePrinterMacGX, VARIANT* ManualDuplexPrint, VARIANT* PrintZoomColumn, VARIANT* PrintZoomRow, VARIANT* PrintZoomPaperWidth, VARIANT* PrintZoomPaperHeight);
};
/////////////////////////////////////////////////////////////////////////////
// Paragraphs wrapper class

class Paragraphs : public COleDispatchDriver
{
public:
	Paragraphs() {}		// Calls COleDispatchDriver default constructor
	Paragraphs(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Paragraphs(const Paragraphs& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	LPUNKNOWN Get_NewEnum();
	long GetCount();
	LPDISPATCH GetFirst();
	LPDISPATCH GetLast();
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	LPDISPATCH GetFormat();
	void SetFormat(LPDISPATCH newValue);
	LPDISPATCH GetTabStops();
	void SetTabStops(LPDISPATCH newValue);
	LPDISPATCH GetBorders();
	void SetBorders(LPDISPATCH newValue);
	VARIANT GetStyle();
	void SetStyle(VARIANT* newValue);
	long GetAlignment();
	void SetAlignment(long nNewValue);
	long GetKeepTogether();
	void SetKeepTogether(long nNewValue);
	long GetKeepWithNext();
	void SetKeepWithNext(long nNewValue);
	long GetPageBreakBefore();
	void SetPageBreakBefore(long nNewValue);
	long GetNoLineNumber();
	void SetNoLineNumber(long nNewValue);
	float GetRightIndent();
	void SetRightIndent(float newValue);
	float GetLeftIndent();
	void SetLeftIndent(float newValue);
	float GetFirstLineIndent();
	void SetFirstLineIndent(float newValue);
	float GetLineSpacing();
	void SetLineSpacing(float newValue);
	long GetLineSpacingRule();
	void SetLineSpacingRule(long nNewValue);
	float GetSpaceBefore();
	void SetSpaceBefore(float newValue);
	float GetSpaceAfter();
	void SetSpaceAfter(float newValue);
	long GetHyphenation();
	void SetHyphenation(long nNewValue);
	long GetWidowControl();
	void SetWidowControl(long nNewValue);
	LPDISPATCH GetShading();
	long GetFarEastLineBreakControl();
	void SetFarEastLineBreakControl(long nNewValue);
	long GetWordWrap();
	void SetWordWrap(long nNewValue);
	long GetHangingPunctuation();
	void SetHangingPunctuation(long nNewValue);
	long GetHalfWidthPunctuationOnTopOfLine();
	void SetHalfWidthPunctuationOnTopOfLine(long nNewValue);
	long GetAddSpaceBetweenFarEastAndAlpha();
	void SetAddSpaceBetweenFarEastAndAlpha(long nNewValue);
	long GetAddSpaceBetweenFarEastAndDigit();
	void SetAddSpaceBetweenFarEastAndDigit(long nNewValue);
	long GetBaseLineAlignment();
	void SetBaseLineAlignment(long nNewValue);
	long GetAutoAdjustRightIndent();
	void SetAutoAdjustRightIndent(long nNewValue);
	long GetDisableLineHeightGrid();
	void SetDisableLineHeightGrid(long nNewValue);
	long GetOutlineLevel();
	void SetOutlineLevel(long nNewValue);
	LPDISPATCH Item(long Index);
	LPDISPATCH Add(VARIANT* Range);
	void CloseUp();
	void OpenUp();
	void OpenOrCloseUp();
	void TabHangingIndent(short Count);
	void TabIndent(short Count);
	void Reset();
	void Space1();
	void Space15();
	void Space2();
	void IndentCharWidth(short Count);
	void IndentFirstLineCharWidth(short Count);
	void OutlinePromote();
	void OutlineDemote();
	void OutlineDemoteToBody();
	void Indent();
	void Outdent();
	float GetCharacterUnitRightIndent();
	void SetCharacterUnitRightIndent(float newValue);
	float GetCharacterUnitLeftIndent();
	void SetCharacterUnitLeftIndent(float newValue);
	float GetCharacterUnitFirstLineIndent();
	void SetCharacterUnitFirstLineIndent(float newValue);
	float GetLineUnitBefore();
	void SetLineUnitBefore(float newValue);
	float GetLineUnitAfter();
	void SetLineUnitAfter(float newValue);
	long GetReadingOrder();
	void SetReadingOrder(long nNewValue);
	long GetSpaceBeforeAuto();
	void SetSpaceBeforeAuto(long nNewValue);
	long GetSpaceAfterAuto();
	void SetSpaceAfterAuto(long nNewValue);
};
/////////////////////////////////////////////////////////////////////////////
// Paragraph wrapper class

class Paragraph : public COleDispatchDriver
{
public:
	Paragraph() {}		// Calls COleDispatchDriver default constructor
	Paragraph(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Paragraph(const Paragraph& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	LPDISPATCH GetRange();
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	LPDISPATCH GetFormat();
	void SetFormat(LPDISPATCH newValue);
	LPDISPATCH GetTabStops();
	void SetTabStops(LPDISPATCH newValue);
	LPDISPATCH GetBorders();
	void SetBorders(LPDISPATCH newValue);
	LPDISPATCH GetDropCap();
	VARIANT GetStyle();
	void SetStyle(VARIANT* newValue);
	long GetAlignment();
	void SetAlignment(long nNewValue);
	long GetKeepTogether();
	void SetKeepTogether(long nNewValue);
	long GetKeepWithNext();
	void SetKeepWithNext(long nNewValue);
	long GetPageBreakBefore();
	void SetPageBreakBefore(long nNewValue);
	long GetNoLineNumber();
	void SetNoLineNumber(long nNewValue);
	float GetRightIndent();
	void SetRightIndent(float newValue);
	float GetLeftIndent();
	void SetLeftIndent(float newValue);
	float GetFirstLineIndent();
	void SetFirstLineIndent(float newValue);
	float GetLineSpacing();
	void SetLineSpacing(float newValue);
	long GetLineSpacingRule();
	void SetLineSpacingRule(long nNewValue);
	float GetSpaceBefore();
	void SetSpaceBefore(float newValue);
	float GetSpaceAfter();
	void SetSpaceAfter(float newValue);
	long GetHyphenation();
	void SetHyphenation(long nNewValue);
	long GetWidowControl();
	void SetWidowControl(long nNewValue);
	LPDISPATCH GetShading();
	long GetFarEastLineBreakControl();
	void SetFarEastLineBreakControl(long nNewValue);
	long GetWordWrap();
	void SetWordWrap(long nNewValue);
	long GetHangingPunctuation();
	void SetHangingPunctuation(long nNewValue);
	long GetHalfWidthPunctuationOnTopOfLine();
	void SetHalfWidthPunctuationOnTopOfLine(long nNewValue);
	long GetAddSpaceBetweenFarEastAndAlpha();
	void SetAddSpaceBetweenFarEastAndAlpha(long nNewValue);
	long GetAddSpaceBetweenFarEastAndDigit();
	void SetAddSpaceBetweenFarEastAndDigit(long nNewValue);
	long GetBaseLineAlignment();
	void SetBaseLineAlignment(long nNewValue);
	long GetAutoAdjustRightIndent();
	void SetAutoAdjustRightIndent(long nNewValue);
	long GetDisableLineHeightGrid();
	void SetDisableLineHeightGrid(long nNewValue);
	long GetOutlineLevel();
	void SetOutlineLevel(long nNewValue);
	void CloseUp();
	void OpenUp();
	void OpenOrCloseUp();
	void TabHangingIndent(short Count);
	void TabIndent(short Count);
	void Reset();
	void Space1();
	void Space15();
	void Space2();
	void IndentCharWidth(short Count);
	void IndentFirstLineCharWidth(short Count);
	LPDISPATCH Next(VARIANT* Count);
	LPDISPATCH Previous(VARIANT* Count);
	void OutlinePromote();
	void OutlineDemote();
	void OutlineDemoteToBody();
	void Indent();
	void Outdent();
	float GetCharacterUnitRightIndent();
	void SetCharacterUnitRightIndent(float newValue);
	float GetCharacterUnitLeftIndent();
	void SetCharacterUnitLeftIndent(float newValue);
	float GetCharacterUnitFirstLineIndent();
	void SetCharacterUnitFirstLineIndent(float newValue);
	float GetLineUnitBefore();
	void SetLineUnitBefore(float newValue);
	float GetLineUnitAfter();
	void SetLineUnitAfter(float newValue);
	long GetReadingOrder();
	void SetReadingOrder(long nNewValue);
	CString GetId();
	void SetId(LPCTSTR lpszNewValue);
	long GetSpaceBeforeAuto();
	void SetSpaceBeforeAuto(long nNewValue);
	long GetSpaceAfterAuto();
	void SetSpaceAfterAuto(long nNewValue);
};
/////////////////////////////////////////////////////////////////////////////
// Selection wrapper class

class Selection : public COleDispatchDriver
{
public:
	Selection() {}		// Calls COleDispatchDriver default constructor
	Selection(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Selection(const Selection& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	CString GetText();
	void SetText(LPCTSTR lpszNewValue);
	LPDISPATCH GetFormattedText();
	void SetFormattedText(LPDISPATCH newValue);
	long GetStart();
	void SetStart(long nNewValue);
	long GetEnd();
	void SetEnd(long nNewValue);
	LPDISPATCH GetFont();
	void SetFont(LPDISPATCH newValue);
	long GetType();
	long GetStoryType();
	VARIANT GetStyle();
	void SetStyle(VARIANT* newValue);
	LPDISPATCH GetTables();
	LPDISPATCH GetWords();
	LPDISPATCH GetSentences();
	LPDISPATCH GetCharacters();
	LPDISPATCH GetFootnotes();
	LPDISPATCH GetEndnotes();
	LPDISPATCH GetComments();
	LPDISPATCH GetCells();
	LPDISPATCH GetSections();
	LPDISPATCH GetParagraphs();
	LPDISPATCH GetBorders();
	void SetBorders(LPDISPATCH newValue);
	LPDISPATCH GetShading();
	LPDISPATCH GetFields();
	LPDISPATCH GetFormFields();
	LPDISPATCH GetFrames();
	LPDISPATCH GetParagraphFormat();
	void SetParagraphFormat(LPDISPATCH newValue);
	LPDISPATCH GetPageSetup();
	void SetPageSetup(LPDISPATCH newValue);
	LPDISPATCH GetBookmarks();
	long GetStoryLength();
	long GetLanguageID();
	void SetLanguageID(long nNewValue);
	long GetLanguageIDFarEast();
	void SetLanguageIDFarEast(long nNewValue);
	long GetLanguageIDOther();
	void SetLanguageIDOther(long nNewValue);
	LPDISPATCH GetHyperlinks();
	LPDISPATCH GetColumns();
	LPDISPATCH GetRows();
	LPDISPATCH GetHeaderFooter();
	BOOL GetIsEndOfRowMark();
	long GetBookmarkID();
	long GetPreviousBookmarkID();
	LPDISPATCH GetFind();
	LPDISPATCH GetRange();
	VARIANT GetInformation(long Type);
	long GetFlags();
	void SetFlags(long nNewValue);
	BOOL GetActive();
	BOOL GetStartIsActive();
	void SetStartIsActive(BOOL bNewValue);
	BOOL GetIPAtEndOfLine();
	BOOL GetExtendMode();
	void SetExtendMode(BOOL bNewValue);
	BOOL GetColumnSelectMode();
	void SetColumnSelectMode(BOOL bNewValue);
	long GetOrientation();
	void SetOrientation(long nNewValue);
	LPDISPATCH GetInlineShapes();
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	LPDISPATCH GetDocument();
	LPDISPATCH GetShapeRange();
	void Select();
	void SetRange(long Start, long End);
	void Collapse(VARIANT* Direction);
	void InsertBefore(LPCTSTR Text);
	void InsertAfter(LPCTSTR Text);
	LPDISPATCH Next(VARIANT* Unit, VARIANT* Count);
	LPDISPATCH Previous(VARIANT* Unit, VARIANT* Count);
	long StartOf(VARIANT* Unit, VARIANT* Extend);
	long EndOf(VARIANT* Unit, VARIANT* Extend);
	long Move(VARIANT* Unit, VARIANT* Count);
	long MoveStart(VARIANT* Unit, VARIANT* Count);
	long MoveEnd(VARIANT* Unit, VARIANT* Count);
	long MoveWhile(VARIANT* Cset, VARIANT* Count);
	long MoveStartWhile(VARIANT* Cset, VARIANT* Count);
	long MoveEndWhile(VARIANT* Cset, VARIANT* Count);
	long MoveUntil(VARIANT* Cset, VARIANT* Count);
	long MoveStartUntil(VARIANT* Cset, VARIANT* Count);
	long MoveEndUntil(VARIANT* Cset, VARIANT* Count);
	void Cut();
	void Copy();
	void Paste();
	void InsertBreak(VARIANT* Type);
	void InsertFile(LPCTSTR FileName, VARIANT* Range, VARIANT* ConfirmConversions, VARIANT* Link, VARIANT* Attachment);
	BOOL InStory(LPDISPATCH Range);
	BOOL InRange(LPDISPATCH Range);
	long Delete(VARIANT* Unit, VARIANT* Count);
	long Expand(VARIANT* Unit);
	void InsertParagraph();
	void InsertParagraphAfter();
	void InsertSymbol(long CharacterNumber, VARIANT* Font, VARIANT* Unicode, VARIANT* Bias);
	void InsertCrossReference(VARIANT* ReferenceType, long ReferenceKind, VARIANT* ReferenceItem, VARIANT* InsertAsHyperlink, VARIANT* IncludePosition);
	void InsertCaption(VARIANT* Label, VARIANT* Title, VARIANT* TitleAutoText, VARIANT* Position);
	void CopyAsPicture();
	void SortAscending();
	void SortDescending();
	BOOL IsEqual(LPDISPATCH Range);
	float Calculate();
	LPDISPATCH GoTo(VARIANT* What, VARIANT* Which, VARIANT* Count, VARIANT* Name);
	LPDISPATCH GoToNext(long What);
	LPDISPATCH GoToPrevious(long What);
	void PasteSpecial(VARIANT* IconIndex, VARIANT* Link, VARIANT* Placement, VARIANT* DisplayAsIcon, VARIANT* DataType, VARIANT* IconFileName, VARIANT* IconLabel);
	LPDISPATCH PreviousField();
	LPDISPATCH NextField();
	void InsertParagraphBefore();
	void InsertCells(VARIANT* ShiftCells);
	void Extend(VARIANT* Character);
	void Shrink();
	long MoveLeft(VARIANT* Unit, VARIANT* Count, VARIANT* Extend);
	long MoveRight(VARIANT* Unit, VARIANT* Count, VARIANT* Extend);
	long MoveUp(VARIANT* Unit, VARIANT* Count, VARIANT* Extend);
	long MoveDown(VARIANT* Unit, VARIANT* Count, VARIANT* Extend);
	long HomeKey(VARIANT* Unit, VARIANT* Extend);
	long EndKey(VARIANT* Unit, VARIANT* Extend);
	void EscapeKey();
	void TypeText(LPCTSTR Text);
	void CopyFormat();
	void PasteFormat();
	void TypeParagraph();
	void TypeBackspace();
	void NextSubdocument();
	void PreviousSubdocument();
	void SelectColumn();
	void SelectCurrentFont();
	void SelectCurrentAlignment();
	void SelectCurrentSpacing();
	void SelectCurrentIndent();
	void SelectCurrentTabs();
	void SelectCurrentColor();
	void CreateTextbox();
	void WholeStory();
	void SelectRow();
	void SplitTable();
	void InsertRows(VARIANT* NumRows);
	void InsertColumns();
	void InsertFormula(VARIANT* Formula, VARIANT* NumberFormat);
	LPDISPATCH NextRevision(VARIANT* Wrap);
	LPDISPATCH PreviousRevision(VARIANT* Wrap);
	void PasteAsNestedTable();
	LPDISPATCH CreateAutoTextEntry(LPCTSTR Name, LPCTSTR StyleName);
	void DetectLanguage();
	void SelectCell();
	void InsertRowsBelow(VARIANT* NumRows);
	void InsertColumnsRight();
	void InsertRowsAbove(VARIANT* NumRows);
	void RtlRun();
	void LtrRun();
	void BoldRun();
	void ItalicRun();
	void RtlPara();
	void LtrPara();
	void InsertDateTime(VARIANT* DateTimeFormat, VARIANT* InsertAsField, VARIANT* InsertAsFullWidth, VARIANT* DateLanguage, VARIANT* CalendarType);
	void Sort(VARIANT* ExcludeHeader, VARIANT* FieldNumber, VARIANT* SortFieldType, VARIANT* SortOrder, VARIANT* FieldNumber2, VARIANT* SortFieldType2, VARIANT* SortOrder2, VARIANT* FieldNumber3, VARIANT* SortFieldType3, VARIANT* SortOrder3, 
		VARIANT* SortColumn, VARIANT* Separator, VARIANT* CaseSensitive, VARIANT* BidiSort, VARIANT* IgnoreThe, VARIANT* IgnoreKashida, VARIANT* IgnoreDiacritics, VARIANT* IgnoreHe, VARIANT* LanguageID);
	LPDISPATCH ConvertToTable(VARIANT* Separator, VARIANT* NumRows, VARIANT* NumColumns, VARIANT* InitialColumnWidth, VARIANT* Format, VARIANT* ApplyBorders, VARIANT* ApplyShading, VARIANT* ApplyFont, VARIANT* ApplyColor, 
		VARIANT* ApplyHeadingRows, VARIANT* ApplyLastRow, VARIANT* ApplyFirstColumn, VARIANT* ApplyLastColumn, VARIANT* AutoFit, VARIANT* AutoFitBehavior, VARIANT* DefaultTableBehavior);
	long GetNoProofing();
	void SetNoProofing(long nNewValue);
	LPDISPATCH GetTopLevelTables();
	BOOL GetLanguageDetected();
	void SetLanguageDetected(BOOL bNewValue);
	float GetFitTextWidth();
	void SetFitTextWidth(float newValue);
};
/////////////////////////////////////////////////////////////////////////////
// InlineShape wrapper class

class InlineShape : public COleDispatchDriver
{
public:
	InlineShape() {}		// Calls COleDispatchDriver default constructor
	InlineShape(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	InlineShape(const InlineShape& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	LPDISPATCH GetBorders();
	void SetBorders(LPDISPATCH newValue);
	LPDISPATCH GetRange();
	LPDISPATCH GetLinkFormat();
	LPDISPATCH GetField();
	LPDISPATCH GetOLEFormat();
	long GetType();
	LPDISPATCH GetHyperlink();
	float GetHeight();
	void SetHeight(float newValue);
	float GetWidth();
	void SetWidth(float newValue);
	float GetScaleHeight();
	void SetScaleHeight(float newValue);
	float GetScaleWidth();
	void SetScaleWidth(float newValue);
	long GetLockAspectRatio();
	void SetLockAspectRatio(long nNewValue);
	LPDISPATCH GetLine();
	LPDISPATCH GetFill();
	LPDISPATCH GetPictureFormat();
	void SetPictureFormat(LPDISPATCH newValue);
	void Activate();
	void Reset();
	void Delete();
	void Select();
	LPDISPATCH ConvertToShape();
	LPDISPATCH GetHorizontalLineFormat();
	LPDISPATCH GetScript();
	LPDISPATCH GetTextEffect();
	void SetTextEffect(LPDISPATCH newValue);
	CString GetAlternativeText();
	void SetAlternativeText(LPCTSTR lpszNewValue);
};
/////////////////////////////////////////////////////////////////////////////
// InlineShapes wrapper class

class InlineShapes : public COleDispatchDriver
{
public:
	InlineShapes() {}		// Calls COleDispatchDriver default constructor
	InlineShapes(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	InlineShapes(const InlineShapes& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	long GetCount();
	LPUNKNOWN Get_NewEnum();
	LPDISPATCH Item(long Index);
	LPDISPATCH AddPicture(LPCTSTR FileName, VARIANT* LinkToFile, VARIANT* SaveWithDocument, VARIANT* Range);
	LPDISPATCH AddOLEObject(VARIANT* ClassType, VARIANT* FileName, VARIANT* LinkToFile, VARIANT* DisplayAsIcon, VARIANT* IconFileName, VARIANT* IconIndex, VARIANT* IconLabel, VARIANT* Range);
	LPDISPATCH AddOLEControl(VARIANT* ClassType, VARIANT* Range);
	LPDISPATCH New(LPDISPATCH Range);
	LPDISPATCH AddHorizontalLine(LPCTSTR FileName, VARIANT* Range);
	LPDISPATCH AddHorizontalLineStandard(VARIANT* Range);
	LPDISPATCH AddPictureBullet(LPCTSTR FileName, VARIANT* Range);
};
/////////////////////////////////////////////////////////////////////////////
// _Font wrapper class

class _Font : public COleDispatchDriver
{
public:
	_Font() {}		// Calls COleDispatchDriver default constructor
	_Font(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	_Font(const _Font& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	LPDISPATCH GetDuplicate();
	long GetBold();
	void SetBold(long nNewValue);
	long GetItalic();
	void SetItalic(long nNewValue);
	long GetHidden();
	void SetHidden(long nNewValue);
	long GetSmallCaps();
	void SetSmallCaps(long nNewValue);
	long GetAllCaps();
	void SetAllCaps(long nNewValue);
	long GetStrikeThrough();
	void SetStrikeThrough(long nNewValue);
	long GetDoubleStrikeThrough();
	void SetDoubleStrikeThrough(long nNewValue);
	long GetColorIndex();
	void SetColorIndex(long nNewValue);
	long GetSubscript();
	void SetSubscript(long nNewValue);
	long GetSuperscript();
	void SetSuperscript(long nNewValue);
	long GetUnderline();
	void SetUnderline(long nNewValue);
	float GetSize();
	void SetSize(float newValue);
	CString GetName();
	void SetName(LPCTSTR lpszNewValue);
	long GetPosition();
	void SetPosition(long nNewValue);
	float GetSpacing();
	void SetSpacing(float newValue);
	long GetScaling();
	void SetScaling(long nNewValue);
	long GetShadow();
	void SetShadow(long nNewValue);
	long GetOutline();
	void SetOutline(long nNewValue);
	long GetEmboss();
	void SetEmboss(long nNewValue);
	float GetKerning();
	void SetKerning(float newValue);
	long GetEngrave();
	void SetEngrave(long nNewValue);
	long GetAnimation();
	void SetAnimation(long nNewValue);
	LPDISPATCH GetBorders();
	void SetBorders(LPDISPATCH newValue);
	LPDISPATCH GetShading();
	long GetEmphasisMark();
	void SetEmphasisMark(long nNewValue);
	BOOL GetDisableCharacterSpaceGrid();
	void SetDisableCharacterSpaceGrid(BOOL bNewValue);
	CString GetNameFarEast();
	void SetNameFarEast(LPCTSTR lpszNewValue);
	CString GetNameAscii();
	void SetNameAscii(LPCTSTR lpszNewValue);
	CString GetNameOther();
	void SetNameOther(LPCTSTR lpszNewValue);
	void Grow();
	void Shrink();
	void Reset();
	void SetAsTemplateDefault();
	long GetColor();
	void SetColor(long nNewValue);
	long GetBoldBi();
	void SetBoldBi(long nNewValue);
	long GetItalicBi();
	void SetItalicBi(long nNewValue);
	float GetSizeBi();
	void SetSizeBi(float newValue);
	CString GetNameBi();
	void SetNameBi(LPCTSTR lpszNewValue);
	long GetColorIndexBi();
	void SetColorIndexBi(long nNewValue);
	long GetDiacriticColor();
	void SetDiacriticColor(long nNewValue);
	long GetUnderlineColor();
	void SetUnderlineColor(long nNewValue);
};
/////////////////////////////////////////////////////////////////////////////
// Table wrapper class

class Table : public COleDispatchDriver
{
public:
	Table() {}		// Calls COleDispatchDriver default constructor
	Table(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Table(const Table& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	LPDISPATCH GetRange();
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	LPDISPATCH GetColumns();
	LPDISPATCH GetRows();
	LPDISPATCH GetBorders();
	void SetBorders(LPDISPATCH newValue);
	LPDISPATCH GetShading();
	BOOL GetUniform();
	long GetAutoFormatType();
	void Select();
	void Delete();
	void SortAscending();
	void SortDescending();
	void AutoFormat(VARIANT* Format, VARIANT* ApplyBorders, VARIANT* ApplyShading, VARIANT* ApplyFont, VARIANT* ApplyColor, VARIANT* ApplyHeadingRows, VARIANT* ApplyLastRow, VARIANT* ApplyFirstColumn, VARIANT* ApplyLastColumn, VARIANT* AutoFit);
	void UpdateAutoFormat();
	LPDISPATCH Cell(long Row, long Column);
	LPDISPATCH Split(VARIANT* BeforeRow);
	LPDISPATCH ConvertToText(VARIANT* Separator, VARIANT* NestedTables);
	void AutoFitBehavior(long Behavior);
	void Sort(VARIANT* ExcludeHeader, VARIANT* FieldNumber, VARIANT* SortFieldType, VARIANT* SortOrder, VARIANT* FieldNumber2, VARIANT* SortFieldType2, VARIANT* SortOrder2, VARIANT* FieldNumber3, VARIANT* SortFieldType3, VARIANT* SortOrder3, 
		VARIANT* CaseSensitive, VARIANT* BidiSort, VARIANT* IgnoreThe, VARIANT* IgnoreKashida, VARIANT* IgnoreDiacritics, VARIANT* IgnoreHe, VARIANT* LanguageID);
	LPDISPATCH GetTables();
	long GetNestingLevel();
	BOOL GetAllowPageBreaks();
	void SetAllowPageBreaks(BOOL bNewValue);
	BOOL GetAllowAutoFit();
	void SetAllowAutoFit(BOOL bNewValue);
	float GetPreferredWidth();
	void SetPreferredWidth(float newValue);
	long GetPreferredWidthType();
	void SetPreferredWidthType(long nNewValue);
	float GetTopPadding();
	void SetTopPadding(float newValue);
	float GetBottomPadding();
	void SetBottomPadding(float newValue);
	float GetLeftPadding();
	void SetLeftPadding(float newValue);
	float GetRightPadding();
	void SetRightPadding(float newValue);
	float GetSpacing();
	void SetSpacing(float newValue);
	long GetTableDirection();
	void SetTableDirection(long nNewValue);
	CString GetId();
	void SetId(LPCTSTR lpszNewValue);
};
/////////////////////////////////////////////////////////////////////////////
// Row wrapper class

class Row : public COleDispatchDriver
{
public:
	Row() {}		// Calls COleDispatchDriver default constructor
	Row(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Row(const Row& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	LPDISPATCH GetRange();
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	long GetAllowBreakAcrossPages();
	void SetAllowBreakAcrossPages(long nNewValue);
	long GetAlignment();
	void SetAlignment(long nNewValue);
	long GetHeadingFormat();
	void SetHeadingFormat(long nNewValue);
	float GetSpaceBetweenColumns();
	void SetSpaceBetweenColumns(float newValue);
	float GetHeight();
	void SetHeight(float newValue);
	long GetHeightRule();
	void SetHeightRule(long nNewValue);
	float GetLeftIndent();
	void SetLeftIndent(float newValue);
	BOOL GetIsLast();
	BOOL GetIsFirst();
	long GetIndex();
	LPDISPATCH GetCells();
	LPDISPATCH GetBorders();
	void SetBorders(LPDISPATCH newValue);
	LPDISPATCH GetShading();
	LPDISPATCH GetNext();
	LPDISPATCH GetPrevious();
	void Select();
	void Delete();
	void SetLeftIndent(float LeftIndent, long RulerStyle);
	void SetHeight(float RowHeight, long HeightRule);
	LPDISPATCH ConvertToText(VARIANT* Separator, VARIANT* NestedTables);
	long GetNestingLevel();
	CString GetId();
	void SetId(LPCTSTR lpszNewValue);
};
/////////////////////////////////////////////////////////////////////////////
// Column wrapper class

class Column : public COleDispatchDriver
{
public:
	Column() {}		// Calls COleDispatchDriver default constructor
	Column(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Column(const Column& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	float GetWidth();
	void SetWidth(float newValue);
	BOOL GetIsFirst();
	BOOL GetIsLast();
	long GetIndex();
	LPDISPATCH GetCells();
	LPDISPATCH GetBorders();
	void SetBorders(LPDISPATCH newValue);
	LPDISPATCH GetShading();
	LPDISPATCH GetNext();
	LPDISPATCH GetPrevious();
	void Select();
	void Delete();
	void SetWidth(float ColumnWidth, long RulerStyle);
	void AutoFit();
	void Sort(VARIANT* ExcludeHeader, VARIANT* SortFieldType, VARIANT* SortOrder, VARIANT* CaseSensitive, VARIANT* BidiSort, VARIANT* IgnoreThe, VARIANT* IgnoreKashida, VARIANT* IgnoreDiacritics, VARIANT* IgnoreHe, VARIANT* LanguageID);
	long GetNestingLevel();
	float GetPreferredWidth();
	void SetPreferredWidth(float newValue);
	long GetPreferredWidthType();
	void SetPreferredWidthType(long nNewValue);
};
/////////////////////////////////////////////////////////////////////////////
// Cell wrapper class

class Cell : public COleDispatchDriver
{
public:
	Cell() {}		// Calls COleDispatchDriver default constructor
	Cell(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Cell(const Cell& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	LPDISPATCH GetRange();
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	long GetRowIndex();
	long GetColumnIndex();
	float GetWidth();
	void SetWidth(float newValue);
	float GetHeight();
	void SetHeight(float newValue);
	long GetHeightRule();
	void SetHeightRule(long nNewValue);
	long GetVerticalAlignment();
	void SetVerticalAlignment(long nNewValue);
	LPDISPATCH GetColumn();
	LPDISPATCH GetRow();
	LPDISPATCH GetNext();
	LPDISPATCH GetPrevious();
	LPDISPATCH GetShading();
	LPDISPATCH GetBorders();
	void SetBorders(LPDISPATCH newValue);
	void Select();
	void Delete(VARIANT* ShiftCells);
	void Formula(VARIANT* Formula, VARIANT* NumFormat);
	void SetWidth(float ColumnWidth, long RulerStyle);
	void SetHeight(VARIANT* RowHeight, long HeightRule);
	void Merge(LPDISPATCH MergeTo);
	void Split(VARIANT* NumRows, VARIANT* NumColumns);
	void AutoSum();
	LPDISPATCH GetTables();
	long GetNestingLevel();
	BOOL GetWordWrap();
	void SetWordWrap(BOOL bNewValue);
	float GetPreferredWidth();
	void SetPreferredWidth(float newValue);
	BOOL GetFitText();
	void SetFitText(BOOL bNewValue);
	float GetTopPadding();
	void SetTopPadding(float newValue);
	float GetBottomPadding();
	void SetBottomPadding(float newValue);
	float GetLeftPadding();
	void SetLeftPadding(float newValue);
	float GetRightPadding();
	void SetRightPadding(float newValue);
	CString GetId();
	void SetId(LPCTSTR lpszNewValue);
	long GetPreferredWidthType();
	void SetPreferredWidthType(long nNewValue);
};
/////////////////////////////////////////////////////////////////////////////
// Tables wrapper class

class Tables : public COleDispatchDriver
{
public:
	Tables() {}		// Calls COleDispatchDriver default constructor
	Tables(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Tables(const Tables& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	LPUNKNOWN Get_NewEnum();
	long GetCount();
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	LPDISPATCH Item(long Index);
	LPDISPATCH Add(LPDISPATCH Range, long NumRows, long NumColumns, VARIANT* DefaultTableBehavior, VARIANT* AutoFitBehavior);
	long GetNestingLevel();
};
/////////////////////////////////////////////////////////////////////////////
// Rows wrapper class

class Rows : public COleDispatchDriver
{
public:
	Rows() {}		// Calls COleDispatchDriver default constructor
	Rows(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Rows(const Rows& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	LPUNKNOWN Get_NewEnum();
	long GetCount();
	long GetAllowBreakAcrossPages();
	void SetAllowBreakAcrossPages(long nNewValue);
	long GetAlignment();
	void SetAlignment(long nNewValue);
	long GetHeadingFormat();
	void SetHeadingFormat(long nNewValue);
	float GetSpaceBetweenColumns();
	void SetSpaceBetweenColumns(float newValue);
	float GetHeight();
	void SetHeight(float newValue);
	long GetHeightRule();
	void SetHeightRule(long nNewValue);
	float GetLeftIndent();
	void SetLeftIndent(float newValue);
	LPDISPATCH GetFirst();
	LPDISPATCH GetLast();
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	LPDISPATCH GetBorders();
	void SetBorders(LPDISPATCH newValue);
	LPDISPATCH GetShading();
	LPDISPATCH Item(long Index);
	LPDISPATCH Add(VARIANT* BeforeRow);
	void Select();
	void Delete();
	void SetLeftIndent(float LeftIndent, long RulerStyle);
	void SetHeight(float RowHeight, long HeightRule);
	void DistributeHeight();
	LPDISPATCH ConvertToText(VARIANT* Separator, VARIANT* NestedTables);
	long GetWrapAroundText();
	void SetWrapAroundText(long nNewValue);
	float GetDistanceTop();
	void SetDistanceTop(float newValue);
	float GetDistanceBottom();
	void SetDistanceBottom(float newValue);
	float GetDistanceLeft();
	void SetDistanceLeft(float newValue);
	float GetDistanceRight();
	void SetDistanceRight(float newValue);
	float GetHorizontalPosition();
	void SetHorizontalPosition(float newValue);
	float GetVerticalPosition();
	void SetVerticalPosition(float newValue);
	long GetRelativeHorizontalPosition();
	void SetRelativeHorizontalPosition(long nNewValue);
	long GetRelativeVerticalPosition();
	void SetRelativeVerticalPosition(long nNewValue);
	long GetAllowOverlap();
	void SetAllowOverlap(long nNewValue);
	long GetNestingLevel();
	long GetTableDirection();
	void SetTableDirection(long nNewValue);
};
/////////////////////////////////////////////////////////////////////////////
// Columns wrapper class

class Columns : public COleDispatchDriver
{
public:
	Columns() {}		// Calls COleDispatchDriver default constructor
	Columns(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Columns(const Columns& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	LPUNKNOWN Get_NewEnum();
	long GetCount();
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	LPDISPATCH GetFirst();
	LPDISPATCH GetLast();
	float GetWidth();
	void SetWidth(float newValue);
	LPDISPATCH GetBorders();
	void SetBorders(LPDISPATCH newValue);
	LPDISPATCH GetShading();
	LPDISPATCH Item(long Index);
	LPDISPATCH Add(VARIANT* BeforeColumn);
	void Select();
	void Delete();
	void SetWidth(float ColumnWidth, long RulerStyle);
	void AutoFit();
	void DistributeWidth();
	long GetNestingLevel();
	float GetPreferredWidth();
	void SetPreferredWidth(float newValue);
	long GetPreferredWidthType();
	void SetPreferredWidthType(long nNewValue);
};
/////////////////////////////////////////////////////////////////////////////
// Cells wrapper class

class Cells : public COleDispatchDriver
{
public:
	Cells() {}		// Calls COleDispatchDriver default constructor
	Cells(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Cells(const Cells& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	LPUNKNOWN Get_NewEnum();
	long GetCount();
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	float GetWidth();
	void SetWidth(float newValue);
	float GetHeight();
	void SetHeight(float newValue);
	long GetHeightRule();
	void SetHeightRule(long nNewValue);
	long GetVerticalAlignment();
	void SetVerticalAlignment(long nNewValue);
	LPDISPATCH GetBorders();
	void SetBorders(LPDISPATCH newValue);
	LPDISPATCH GetShading();
	LPDISPATCH Item(long Index);
	LPDISPATCH Add(VARIANT* BeforeCell);
	void Delete(VARIANT* ShiftCells);
	void SetWidth(float ColumnWidth, long RulerStyle);
	void SetHeight(VARIANT* RowHeight, long HeightRule);
	void Merge();
	void Split(VARIANT* NumRows, VARIANT* NumColumns, VARIANT* MergeBeforeSplit);
	void DistributeHeight();
	void DistributeWidth();
	void AutoFit();
	long GetNestingLevel();
	float GetPreferredWidth();
	void SetPreferredWidth(float newValue);
	long GetPreferredWidthType();
	void SetPreferredWidthType(long nNewValue);
};
/////////////////////////////////////////////////////////////////////////////
// Borders wrapper class

class Borders : public COleDispatchDriver
{
public:
	Borders() {}		// Calls COleDispatchDriver default constructor
	Borders(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Borders(const Borders& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	LPUNKNOWN Get_NewEnum();
	long GetCount();
	long GetEnable();
	void SetEnable(long nNewValue);
	long GetDistanceFromTop();
	void SetDistanceFromTop(long nNewValue);
	BOOL GetShadow();
	void SetShadow(BOOL bNewValue);
	long GetInsideLineStyle();
	void SetInsideLineStyle(long nNewValue);
	long GetOutsideLineStyle();
	void SetOutsideLineStyle(long nNewValue);
	long GetInsideLineWidth();
	void SetInsideLineWidth(long nNewValue);
	long GetOutsideLineWidth();
	void SetOutsideLineWidth(long nNewValue);
	long GetInsideColorIndex();
	void SetInsideColorIndex(long nNewValue);
	long GetOutsideColorIndex();
	void SetOutsideColorIndex(long nNewValue);
	long GetDistanceFromLeft();
	void SetDistanceFromLeft(long nNewValue);
	long GetDistanceFromBottom();
	void SetDistanceFromBottom(long nNewValue);
	long GetDistanceFromRight();
	void SetDistanceFromRight(long nNewValue);
	BOOL GetAlwaysInFront();
	void SetAlwaysInFront(BOOL bNewValue);
	BOOL GetSurroundHeader();
	void SetSurroundHeader(BOOL bNewValue);
	BOOL GetSurroundFooter();
	void SetSurroundFooter(BOOL bNewValue);
	BOOL GetJoinBorders();
	void SetJoinBorders(BOOL bNewValue);
	BOOL GetHasHorizontal();
	BOOL GetHasVertical();
	long GetDistanceFrom();
	void SetDistanceFrom(long nNewValue);
	BOOL GetEnableFirstPageInSection();
	void SetEnableFirstPageInSection(BOOL bNewValue);
	BOOL GetEnableOtherPagesInSection();
	void SetEnableOtherPagesInSection(BOOL bNewValue);
	LPDISPATCH Item(long Index);
	void ApplyPageBordersToAllSections();
	long GetInsideColor();
	void SetInsideColor(long nNewValue);
	long GetOutsideColor();
	void SetOutsideColor(long nNewValue);
};
/////////////////////////////////////////////////////////////////////////////
// Border wrapper class

class Border : public COleDispatchDriver
{
public:
	Border() {}		// Calls COleDispatchDriver default constructor
	Border(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Border(const Border& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	BOOL GetVisible();
	void SetVisible(BOOL bNewValue);
	long GetColorIndex();
	void SetColorIndex(long nNewValue);
	BOOL GetInside();
	long GetLineStyle();
	void SetLineStyle(long nNewValue);
	long GetLineWidth();
	void SetLineWidth(long nNewValue);
	long GetArtStyle();
	void SetArtStyle(long nNewValue);
	long GetArtWidth();
	void SetArtWidth(long nNewValue);
	long GetColor();
	void SetColor(long nNewValue);
};
/////////////////////////////////////////////////////////////////////////////
// Shading wrapper class

class Shading : public COleDispatchDriver
{
public:
	Shading() {}		// Calls COleDispatchDriver default constructor
	Shading(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Shading(const Shading& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	long GetForegroundPatternColorIndex();
	void SetForegroundPatternColorIndex(long nNewValue);
	long GetBackgroundPatternColorIndex();
	void SetBackgroundPatternColorIndex(long nNewValue);
	long GetTexture();
	void SetTexture(long nNewValue);
	long GetForegroundPatternColor();
	void SetForegroundPatternColor(long nNewValue);
	long GetBackgroundPatternColor();
	void SetBackgroundPatternColor(long nNewValue);
};
/////////////////////////////////////////////////////////////////////////////
// Range wrapper class

class Range : public COleDispatchDriver
{
public:
	Range() {}		// Calls COleDispatchDriver default constructor
	Range(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	Range(const Range& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	CString GetText();
	void SetText(LPCTSTR lpszNewValue);
	LPDISPATCH GetFormattedText();
	void SetFormattedText(LPDISPATCH newValue);
	long GetStart();
	void SetStart(long nNewValue);
	long GetEnd();
	void SetEnd(long nNewValue);
	LPDISPATCH GetFont();
	void SetFont(LPDISPATCH newValue);
	LPDISPATCH GetDuplicate();
	long GetStoryType();
	LPDISPATCH GetTables();
	LPDISPATCH GetWords();
	LPDISPATCH GetSentences();
	LPDISPATCH GetCharacters();
	LPDISPATCH GetFootnotes();
	LPDISPATCH GetEndnotes();
	LPDISPATCH GetComments();
	LPDISPATCH GetCells();
	LPDISPATCH GetSections();
	LPDISPATCH GetParagraphs();
	LPDISPATCH GetBorders();
	void SetBorders(LPDISPATCH newValue);
	LPDISPATCH GetShading();
	LPDISPATCH GetTextRetrievalMode();
	void SetTextRetrievalMode(LPDISPATCH newValue);
	LPDISPATCH GetFields();
	LPDISPATCH GetFormFields();
	LPDISPATCH GetFrames();
	LPDISPATCH GetParagraphFormat();
	void SetParagraphFormat(LPDISPATCH newValue);
	LPDISPATCH GetListFormat();
	LPDISPATCH GetBookmarks();
	LPDISPATCH GetApplication();
	long GetCreator();
	LPDISPATCH GetParent();
	long GetBold();
	void SetBold(long nNewValue);
	long GetItalic();
	void SetItalic(long nNewValue);
	long GetUnderline();
	void SetUnderline(long nNewValue);
	long GetEmphasisMark();
	void SetEmphasisMark(long nNewValue);
	BOOL GetDisableCharacterSpaceGrid();
	void SetDisableCharacterSpaceGrid(BOOL bNewValue);
	LPDISPATCH GetRevisions();
	VARIANT GetStyle();
	void SetStyle(VARIANT* newValue);
	long GetStoryLength();
	long GetLanguageID();
	void SetLanguageID(long nNewValue);
	LPDISPATCH GetSynonymInfo();
	LPDISPATCH GetHyperlinks();
	LPDISPATCH GetListParagraphs();
	LPDISPATCH GetSubdocuments();
	BOOL GetGrammarChecked();
	void SetGrammarChecked(BOOL bNewValue);
	BOOL GetSpellingChecked();
	void SetSpellingChecked(BOOL bNewValue);
	long GetHighlightColorIndex();
	void SetHighlightColorIndex(long nNewValue);
	LPDISPATCH GetColumns();
	LPDISPATCH GetRows();
	BOOL GetIsEndOfRowMark();
	long GetBookmarkID();
	long GetPreviousBookmarkID();
	LPDISPATCH GetFind();
	LPDISPATCH GetPageSetup();
	void SetPageSetup(LPDISPATCH newValue);
	LPDISPATCH GetShapeRange();
	long GetCase();
	void SetCase(long nNewValue);
	VARIANT GetInformation(long Type);
	LPDISPATCH GetReadabilityStatistics();
	LPDISPATCH GetGrammaticalErrors();
	LPDISPATCH GetSpellingErrors();
	long GetOrientation();
	void SetOrientation(long nNewValue);
	LPDISPATCH GetInlineShapes();
	LPDISPATCH GetNextStoryRange();
	long GetLanguageIDFarEast();
	void SetLanguageIDFarEast(long nNewValue);
	long GetLanguageIDOther();
	void SetLanguageIDOther(long nNewValue);
	void Select();
	void SetRange(long Start, long End);
	void Collapse(VARIANT* Direction);
	void InsertBefore(LPCTSTR Text);
	void InsertAfter(LPCTSTR Text);
	LPDISPATCH Next(VARIANT* Unit, VARIANT* Count);
	LPDISPATCH Previous(VARIANT* Unit, VARIANT* Count);
	long StartOf(VARIANT* Unit, VARIANT* Extend);
	long EndOf(VARIANT* Unit, VARIANT* Extend);
	long Move(VARIANT* Unit, VARIANT* Count);
	long MoveStart(VARIANT* Unit, VARIANT* Count);
	long MoveEnd(VARIANT* Unit, VARIANT* Count);
	long MoveWhile(VARIANT* Cset, VARIANT* Count);
	long MoveStartWhile(VARIANT* Cset, VARIANT* Count);
	long MoveEndWhile(VARIANT* Cset, VARIANT* Count);
	long MoveUntil(VARIANT* Cset, VARIANT* Count);
	long MoveStartUntil(VARIANT* Cset, VARIANT* Count);
	long MoveEndUntil(VARIANT* Cset, VARIANT* Count);
	void Cut();
	void Copy();
	void Paste();
	void InsertBreak(VARIANT* Type);
	void InsertFile(LPCTSTR FileName, VARIANT* Range, VARIANT* ConfirmConversions, VARIANT* Link, VARIANT* Attachment);
	BOOL InStory(LPDISPATCH Range);
	BOOL InRange(LPDISPATCH Range);
	long Delete(VARIANT* Unit, VARIANT* Count);
	void WholeStory();
	long Expand(VARIANT* Unit);
	void InsertParagraph();
	void InsertParagraphAfter();
	void InsertSymbol(long CharacterNumber, VARIANT* Font, VARIANT* Unicode, VARIANT* Bias);
	void InsertCrossReference(VARIANT* ReferenceType, long ReferenceKind, VARIANT* ReferenceItem, VARIANT* InsertAsHyperlink, VARIANT* IncludePosition);
	void InsertCaption(VARIANT* Label, VARIANT* Title, VARIANT* TitleAutoText, VARIANT* Position);
	void CopyAsPicture();
	void SortAscending();
	void SortDescending();
	BOOL IsEqual(LPDISPATCH Range);
	float Calculate();
	LPDISPATCH GoTo(VARIANT* What, VARIANT* Which, VARIANT* Count, VARIANT* Name);
	LPDISPATCH GoToNext(long What);
	LPDISPATCH GoToPrevious(long What);
	void PasteSpecial(VARIANT* IconIndex, VARIANT* Link, VARIANT* Placement, VARIANT* DisplayAsIcon, VARIANT* DataType, VARIANT* IconFileName, VARIANT* IconLabel);
	void LookupNameProperties();
	long ComputeStatistics(long Statistic);
	void Relocate(long Direction);
	void CheckSynonyms();
	void SubscribeTo(LPCTSTR Edition, VARIANT* Format);
	void CreatePublisher(VARIANT* Edition, VARIANT* ContainsPICT, VARIANT* ContainsRTF, VARIANT* ContainsText);
	void InsertAutoText();
	void InsertDatabase(VARIANT* Format, VARIANT* Style, VARIANT* LinkToSource, VARIANT* Connection, VARIANT* SQLStatement, VARIANT* SQLStatement1, VARIANT* PasswordDocument, VARIANT* PasswordTemplate, VARIANT* WritePasswordDocument, 
		VARIANT* WritePasswordTemplate, VARIANT* DataSource, VARIANT* From, VARIANT* To, VARIANT* IncludeFields);
	void AutoFormat();
	void CheckGrammar();
	void CheckSpelling(VARIANT* CustomDictionary, VARIANT* IgnoreUppercase, VARIANT* AlwaysSuggest, VARIANT* CustomDictionary2, VARIANT* CustomDictionary3, VARIANT* CustomDictionary4, VARIANT* CustomDictionary5, VARIANT* CustomDictionary6, 
		VARIANT* CustomDictionary7, VARIANT* CustomDictionary8, VARIANT* CustomDictionary9, VARIANT* CustomDictionary10);
	LPDISPATCH GetSpellingSuggestions(VARIANT* CustomDictionary, VARIANT* IgnoreUppercase, VARIANT* MainDictionary, VARIANT* SuggestionMode, VARIANT* CustomDictionary2, VARIANT* CustomDictionary3, VARIANT* CustomDictionary4, 
		VARIANT* CustomDictionary5, VARIANT* CustomDictionary6, VARIANT* CustomDictionary7, VARIANT* CustomDictionary8, VARIANT* CustomDictionary9, VARIANT* CustomDictionary10);
	void InsertParagraphBefore();
	void NextSubdocument();
	void PreviousSubdocument();
	void ConvertHangulAndHanja(VARIANT* ConversionsMode, VARIANT* FastConversion, VARIANT* CheckHangulEnding, VARIANT* EnableRecentOrdering, VARIANT* CustomDictionary);
	void PasteAsNestedTable();
	void ModifyEnclosure(VARIANT* Style, VARIANT* Symbol, VARIANT* EnclosedText);
	void PhoneticGuide(LPCTSTR Text, long Alignment, long Raise, long FontSize, LPCTSTR FontName);
	void InsertDateTime(VARIANT* DateTimeFormat, VARIANT* InsertAsField, VARIANT* InsertAsFullWidth, VARIANT* DateLanguage, VARIANT* CalendarType);
	void Sort(VARIANT* ExcludeHeader, VARIANT* FieldNumber, VARIANT* SortFieldType, VARIANT* SortOrder, VARIANT* FieldNumber2, VARIANT* SortFieldType2, VARIANT* SortOrder2, VARIANT* FieldNumber3, VARIANT* SortFieldType3, VARIANT* SortOrder3, 
		VARIANT* SortColumn, VARIANT* Separator, VARIANT* CaseSensitive, VARIANT* BidiSort, VARIANT* IgnoreThe, VARIANT* IgnoreKashida, VARIANT* IgnoreDiacritics, VARIANT* IgnoreHe, VARIANT* LanguageID);
	void DetectLanguage();
	LPDISPATCH ConvertToTable(VARIANT* Separator, VARIANT* NumRows, VARIANT* NumColumns, VARIANT* InitialColumnWidth, VARIANT* Format, VARIANT* ApplyBorders, VARIANT* ApplyShading, VARIANT* ApplyFont, VARIANT* ApplyColor, 
		VARIANT* ApplyHeadingRows, VARIANT* ApplyLastRow, VARIANT* ApplyFirstColumn, VARIANT* ApplyLastColumn, VARIANT* AutoFit, VARIANT* AutoFitBehavior, VARIANT* DefaultTableBehavior);
	void TCSCConverter(long WdTCSCConverterDirection, BOOL CommonTerms, BOOL UseVariants);
	BOOL GetLanguageDetected();
	void SetLanguageDetected(BOOL bNewValue);
	float GetFitTextWidth();
	void SetFitTextWidth(float newValue);
	long GetHorizontalInVertical();
	void SetHorizontalInVertical(long nNewValue);
	long GetTwoLinesInOne();
	void SetTwoLinesInOne(long nNewValue);
	BOOL GetCombineCharacters();
	void SetCombineCharacters(BOOL bNewValue);
	long GetNoProofing();
	void SetNoProofing(long nNewValue);
	LPDISPATCH GetTopLevelTables();
	LPDISPATCH GetScripts();
	long GetCharacterWidth();
	void SetCharacterWidth(long nNewValue);
	long GetKana();
	void SetKana(long nNewValue);
	long GetBoldBi();
	void SetBoldBi(long nNewValue);
	long GetItalicBi();
	void SetItalicBi(long nNewValue);
	CString GetId();
	void SetId(LPCTSTR lpszNewValue);
};

// Machine generated IDispatch wrapper class(es) created with Add Class from Typelib Wizard

// CFind wrapper class

class CFind : public COleDispatchDriver
{
public:
	CFind(){} // Calls COleDispatchDriver default constructor
	CFind(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CFind(const CFind& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// Attributes
public:

	// Operations
public:


	// Find methods
public:
	LPDISPATCH get_Application()
	{
		LPDISPATCH result;
		InvokeHelper(0x3e8, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_Creator()
	{
		long result;
		InvokeHelper(0x3e9, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Parent()
	{
		LPDISPATCH result;
		InvokeHelper(0x3ea, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	BOOL get_Forward()
	{
		BOOL result;
		InvokeHelper(0xa, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_Forward(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0xa, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_Font()
	{
		LPDISPATCH result;
		InvokeHelper(0xb, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	void put_Font(LPDISPATCH newValue)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0xb, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_Found()
	{
		BOOL result;
		InvokeHelper(0xc, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	BOOL get_MatchAllWordForms()
	{
		BOOL result;
		InvokeHelper(0xd, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_MatchAllWordForms(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0xd, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_MatchCase()
	{
		BOOL result;
		InvokeHelper(0xe, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_MatchCase(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0xe, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_MatchWildcards()
	{
		BOOL result;
		InvokeHelper(0xf, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_MatchWildcards(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0xf, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_MatchSoundsLike()
	{
		BOOL result;
		InvokeHelper(0x10, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_MatchSoundsLike(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x10, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_MatchWholeWord()
	{
		BOOL result;
		InvokeHelper(0x11, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_MatchWholeWord(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x11, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_MatchFuzzy()
	{
		BOOL result;
		InvokeHelper(0x28, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_MatchFuzzy(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x28, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_MatchByte()
	{
		BOOL result;
		InvokeHelper(0x29, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_MatchByte(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x29, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_ParagraphFormat()
	{
		LPDISPATCH result;
		InvokeHelper(0x12, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	void put_ParagraphFormat(LPDISPATCH newValue)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x12, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	VARIANT get_Style()
	{
		VARIANT result;
		InvokeHelper(0x13, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
		return result;
	}
	void put_Style(VARIANT * newValue)
	{
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x13, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	CString get_Text()
	{
		CString result;
		InvokeHelper(0x16, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	void put_Text(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x16, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_LanguageID()
	{
		long result;
		InvokeHelper(0x17, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_LanguageID(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x17, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_Highlight()
	{
		long result;
		InvokeHelper(0x18, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_Highlight(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x18, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_Replacement()
	{
		LPDISPATCH result;
		InvokeHelper(0x19, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Frame()
	{
		LPDISPATCH result;
		InvokeHelper(0x1a, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_Wrap()
	{
		long result;
		InvokeHelper(0x1b, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_Wrap(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x1b, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_Format()
	{
		BOOL result;
		InvokeHelper(0x1c, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_Format(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x1c, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_LanguageIDFarEast()
	{
		long result;
		InvokeHelper(0x1d, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_LanguageIDFarEast(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x1d, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_LanguageIDOther()
	{
		long result;
		InvokeHelper(0x3c, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_LanguageIDOther(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x3c, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_CorrectHangulEndings()
	{
		BOOL result;
		InvokeHelper(0x3d, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_CorrectHangulEndings(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x3d, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL ExecuteOld(VARIANT * FindText, VARIANT * MatchCase, VARIANT * MatchWholeWord, VARIANT * MatchWildcards, VARIANT * MatchSoundsLike, VARIANT * MatchAllWordForms, VARIANT * Forward, VARIANT * Wrap, VARIANT * Format, VARIANT * ReplaceWith, VARIANT * Replace)
	{
		BOOL result;
		static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x1e, DISPATCH_METHOD, VT_BOOL, (void*)&result, parms, FindText, MatchCase, MatchWholeWord, MatchWildcards, MatchSoundsLike, MatchAllWordForms, Forward, Wrap, Format, ReplaceWith, Replace);
		return result;
	}
	void ClearFormatting()
	{
		InvokeHelper(0x1f, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void SetAllFuzzyOptions()
	{
		InvokeHelper(0x20, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void ClearAllFuzzyOptions()
	{
		InvokeHelper(0x21, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	BOOL Execute(VARIANT * FindText, VARIANT * MatchCase, VARIANT * MatchWholeWord, VARIANT * MatchWildcards, VARIANT * MatchSoundsLike, VARIANT * MatchAllWordForms, VARIANT * Forward, VARIANT * Wrap, VARIANT * Format, VARIANT * ReplaceWith, VARIANT * Replace, VARIANT * MatchKashida, VARIANT * MatchDiacritics, VARIANT * MatchAlefHamza, VARIANT * MatchControl)
	{
		BOOL result;
		static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x1bc, DISPATCH_METHOD, VT_BOOL, (void*)&result, parms, FindText, MatchCase, MatchWholeWord, MatchWildcards, MatchSoundsLike, MatchAllWordForms, Forward, Wrap, Format, ReplaceWith, Replace, MatchKashida, MatchDiacritics, MatchAlefHamza, MatchControl);
		return result;
	}
	long get_NoProofing()
	{
		long result;
		InvokeHelper(0x22, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_NoProofing(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x22, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_MatchKashida()
	{
		BOOL result;
		InvokeHelper(0x64, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_MatchKashida(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x64, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_MatchDiacritics()
	{
		BOOL result;
		InvokeHelper(0x65, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_MatchDiacritics(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x65, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_MatchAlefHamza()
	{
		BOOL result;
		InvokeHelper(0x66, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_MatchAlefHamza(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x66, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_MatchControl()
	{
		BOOL result;
		InvokeHelper(0x67, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_MatchControl(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x67, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}

	// Find properties
public:

};


#endif //#ifndef msword9_h
