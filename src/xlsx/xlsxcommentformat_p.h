#ifndef xlsxcommentformat_p_h__
#define xlsxcommentformat_p_h__
#include "xlsxcommentformat.h"
#include "xlsxglobal.h"
#include <QSizeF>
#include <QColor>
QT_BEGIN_NAMESPACE_XLSX
class XLSX_AUTOTEST_EXPORT CommentFormatPrivate
{
private:
    Q_DECLARE_PUBLIC(CommentFormat)
    CommentFormat* q_ptr;
    CommentFormatPrivate(const CommentFormatPrivate& other);
public:
    CommentFormatPrivate& operator=(const CommentFormatPrivate& other);
    QColor m_backgroundColor;
    bool m_showShadow;
    QColor m_shadowColor;
    double m_hOffset;
    double m_vOffset;
    QSizeF m_commentSize;
    Qt::Alignment m_textAlign;
    CommentFormatPrivate(CommentFormat* q);
    CommentFormatPrivate(CommentFormat* q, const CommentFormatPrivate& other);
    virtual ~CommentFormatPrivate();
};









QT_END_NAMESPACE_XLSX
#endif // xlsxcommentformat_p_h__