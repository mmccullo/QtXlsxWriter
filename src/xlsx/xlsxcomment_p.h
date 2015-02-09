#ifndef xlsxcomment_p_h__
#define xlsxcomment_p_h__
#include "xlsxcomment.h"
#include "xlsxcommentformat.h"
QT_BEGIN_NAMESPACE_XLSX
class XLSX_AUTOTEST_EXPORT CommentPrivate
{
private:
    Q_DECLARE_PUBLIC(Comment)
    Comment* q_ptr;
public:
    CommentPrivate(Comment* q);
    CommentPrivate(Comment* q,const QString& auth, const RichString& txt);
    CommentPrivate(Comment* q, const CommentPrivate& other);
    QString m_Author;
    RichString m_Text;
    CommentFormat m_Format;
};
QT_END_NAMESPACE_XLSX
#endif // xlsxcomment_p_h__
