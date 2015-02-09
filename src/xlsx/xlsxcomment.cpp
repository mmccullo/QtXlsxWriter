#include "xlsxcomment.h"
#include "xlsxcomment_p.h"
#ifndef QT_NO_DEBUG_STREAM
#include <QDebug>
#endif
QT_BEGIN_NAMESPACE_XLSX
CommentPrivate::CommentPrivate(Comment* q)
    :q_ptr(q)
{}
CommentPrivate::CommentPrivate(Comment* q, const QString& auth, const RichString& txt)
    : q_ptr(q)
    , m_Author(auth)
    , m_Text(txt)
{}
CommentPrivate::CommentPrivate(Comment* q, const CommentPrivate& other)
    : q_ptr(q)
    , m_Author(other.m_Author)
    , m_Text(other.m_Text)
    , m_Format(other.m_Format)
{}
Comment::Comment(const QString& auth, const RichString& txt)
:d_ptr(new CommentPrivate(this, auth, txt))
{}
Comment::Comment(const Comment &other)
    : d_ptr(new CommentPrivate(this, *(other.d_ptr)))
{}

Comment::Comment() 
    : d_ptr(new CommentPrivate(this))
{}

Comment& Comment::operator=(const Comment &other)
{
    Q_D(Comment);
    d->m_Author = other.author();
    d->m_Text = other.text();
    d->m_Format = other.format();
    return *this;
}

const QString& Comment::author() const
{
    Q_D(const Comment);
    return d->m_Author;
}

const RichString& Comment::text() const
{
    Q_D(const Comment);
    return d->m_Text;
}
void Comment::setAuthor(const QString& auth)
{
    Q_D(Comment);
    d->m_Author = auth;
}

void Comment::setText(const RichString& txt)
{
    Q_D(Comment);
    d->m_Text = txt;
}
Comment::~Comment() {
    if (d_ptr)
        delete d_ptr;
}

const CommentFormat& Comment::format() const
{
    Q_D(const Comment);
    return d->m_Format;
}

void Comment::setFormat(const CommentFormat& val)
{
    Q_D(Comment);
    d->m_Format = val;
}

bool operator==(const Comment &rs1, const Comment &rs2)
{
    return
        rs1.author().compare(rs2.author(), Qt::CaseInsensitive) == 0
        && rs1.text() == rs2.text()
        ;
}

bool operator!=(const Comment &rs1, const Comment &rs2)
{
    return !(rs1 == rs2);
}
#ifndef QT_NO_DEBUG_STREAM
Q_XLSX_EXPORT QDebug operator<<(QDebug dbg, const Comment &rs)
{
    return dbg << "Comment Author: " + rs.author() + " Comment Text: " + rs.text().toHtml();
}
#endif
QT_END_NAMESPACE_XLSX
