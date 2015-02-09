#ifndef xlsxcommentformat_h__
#define xlsxcommentformat_h__
#include "xlsxglobal.h"
#include <QSizeF>
#include <QColor>
QT_BEGIN_NAMESPACE_XLSX
class CommentFormatPrivate;
class Q_XLSX_EXPORT CommentFormat
{
private:
    Q_DECLARE_PRIVATE(CommentFormat)
    CommentFormatPrivate* d_ptr;
public:
    CommentFormat();
    CommentFormat(const CommentFormat& other);
    CommentFormat& operator=(const CommentFormat& other);
    virtual ~CommentFormat();
    const QSizeF& commentSize() const;
    void setCommentSize(const QSizeF& val);
    const double& hOffset() const;
    void setHOffset(const double& val);
    const double& vOffset() const;
    void setVOffset(const double& val);
    void setOffset(const double& h, const double& v);
    const QColor& shadowColor() const;
    void shadowColor(const QColor& val);
    const bool& showShadow() const;
    void setShowShadow(const bool& val);
    const QColor& backgroundColor() const;
    void setBackgroundColor(const QColor& val);
    const Qt::Alignment& textAlign() const;
    void setTextAlign(const Qt::Alignment& val);

};
QT_END_NAMESPACE_XLSX
#endif // xlsxcommentformat_h__