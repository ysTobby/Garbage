#pragma once
class COperateWord
{
public:
	COperateWord(const CString& temppath, const CString& fname) :
		path(temppath),
		Filename(fname),
		success(true),
		wordApp(),docs(),docx(),bookmarks(),bookmark(),range()
	{
		ReplacePath();//path重新替换成正确的
		
		CoInitialize(NULL);

		covZero = COleVariant((short)0);
		covTrue = COleVariant((short)TRUE);
		covFalse = COleVariant((short)FALSE);
		covOptional = COleVariant((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
		covDocxType = COleVariant((short)0);
		start_line = COleVariant();
		end_line = COleVariant();
		dot = COleVariant(path);

		if (!wordApp.CreateDispatch(_T("Word.Application")))
		{
			success = false;
		}

		wordApp.put_Visible(TRUE);
		docs = wordApp.get_Documents();
		docx = docs.Add(dot,covOptional,covOptional,covOptional);
		bookmarks = docx.get_Bookmarks();
	}

	~COperateWord() 
	{
		range.ReleaseDispatch();
		bookmark.ReleaseDispatch();
		bookmarks.ReleaseDispatch();
		docx.ReleaseDispatch();
		docs.ReleaseDispatch();
		wordApp.ReleaseDispatch();
	}
	bool IsSuccess() { return success; }
	void ReplaceBookmark(const CString& bookname, const CString& content, int num);
private:
	void ReplacePath();//替换名称

	CApplication wordApp;
	CDocuments docs;
	CDocument0 docx;
	CBookmarks bookmarks;
	CBookmark0 bookmark;
	CRange range;
	
	bool success;
	CString path;
	CString Filename;
	COleVariant covZero, covTrue, covFalse, covOptional, covDocxType,
		start_line, end_line, dot;
};

