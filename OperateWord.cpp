#include "stdafx.h"
#include "OperateWord.h"

void COperateWord::ReplaceBookmark(const CString& bookname, const CString& content, int num)
{
	//替换书签的内容
	CString temps;
	for (int i = 1; i <= num; i++)
	{
		temps.Format(_T("%d"), i);
		temps = bookname + temps;

		bookmark = bookmarks.Item(&_variant_t(temps));
		range = bookmark.get_Range();
		range.put_Text(content);
	}
}

void COperateWord::ReplacePath()
{
	path.Replace(_T("\\peijin.exe"), Filename);
}
