#ifndef xlsxspreadsheetcomments_h__
#define xlsxspreadsheetcomments_h__
#include "xlsxabstractooxmlfile.h"
#include "xlsxcomment.h"
QT_BEGIN_NAMESPACE_XLSX
class CellReference;
class SpreadSheetCommentPrivate;
class Worksheet;
class WorksheetPrivate;
class XLSX_AUTOTEST_EXPORT SpreadSheetComment : public AbstractOOXmlFile
{
    Q_DECLARE_PRIVATE(SpreadSheetComment)
public:
    ~SpreadSheetComment();
    Comment getComment(int row, int col) const;
    void setComment(int row, int col, const Comment& val);
    void removeComment(int row, int col);
    void moveComment(int fromRow, int fromCol, int toRow, int toCol);
    void setComment(const CellReference& rowCol, const Comment& val);
    void removeComment(const CellReference& rowCol);
    void moveComment(const CellReference& fromRowCol, int toRow, int toCol);
    void moveComment(const CellReference& fromRowCol, const CellReference& toRowCol);
    void moveComment(int fromRow, int fromCol, const CellReference& toRowCol);
    virtual void saveToXmlFile(QIODevice *device) const;
    virtual bool loadFromXmlFile(QIODevice *device);
    void saveShapeToXmlFile(QIODevice *device) const;
    QByteArray saveShapeToXmlData()const;
    bool hasComment(int row, int col) const;
    bool hasComment(const CellReference& rowCol) const;
    bool isEmpty() const;
private:
    bool hasRow(int row) const;
    SpreadSheetComment(Worksheet* parShe,CreateFlag flag);
    SpreadSheetComment(const SpreadSheetComment& other);
    //same as SharedStrings::writeRichStringPart_rPr that method should be either global or into RichString
    void writeRichStringPart_rPr(QXmlStreamWriter &writer, const Format &format) const;
    friend class Worksheet;
    friend class WorksheetPrivate;
};

QT_END_NAMESPACE_XLSX
#endif // xlsxspreadsheetcomments_h__