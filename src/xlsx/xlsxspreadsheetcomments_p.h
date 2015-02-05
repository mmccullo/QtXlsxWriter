#ifndef xlsxspreadsheetcomments_p_h__
#define xlsxspreadsheetcomments_p_h__
#include "xlsxglobal.h"
#include "xlsxcomment.h"
#include "xlsxspreadsheetcomments.h"
#include "xlsxabstractooxmlfile_p.h"
#include <QMap>
QT_BEGIN_NAMESPACE_XLSX
class XLSX_AUTOTEST_EXPORT SpreadSheetCommentPrivate : public AbstractOOXmlFilePrivate {
	Q_DECLARE_PUBLIC(SpreadSheetComment)
private:
	QMap<int, QMap<int, Comment* >* > m_Comments;
public:
	SpreadSheetCommentPrivate(SpreadSheetComment* p, SpreadSheetComment::CreateFlag flag);
	virtual ~SpreadSheetCommentPrivate();
	const Comment* getComment(int row, int col)const;
	void setComment(int row, int col, const Comment& val);
	void removeComment(int row, int col);
	void moveComment(int fromRow, int fromCol, int toRow, int toCol);
	bool hasRow(int row) const;
	bool isEmpty() const;
	const decltype(m_Comments)& comments() const;
};



QT_END_NAMESPACE_XLSX
#endif // xlsxspreadsheetcomments_p_h__