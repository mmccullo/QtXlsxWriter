#include "xlsxcommentformat_p.h"
#include "xlsxcommentformat.h"
#include <QDebug>
QT_BEGIN_NAMESPACE_XLSX
CommentFormatPrivate::CommentFormatPrivate(CommentFormat* q)
    : q_ptr(q)
    , m_hOffset(11.25)
    , m_vOffset(-7.5)
    , m_backgroundColor("#ffffe1")
    , m_showShadow(true)
    , m_shadowColor(Qt::black)
    , m_commentSize(108.0,59.25)
    , m_textAlign(Qt::AlignLeft | Qt::AlignTop)
{}
CommentFormatPrivate::CommentFormatPrivate(CommentFormat* q, const CommentFormatPrivate& other)
    : q_ptr(q)
    , m_hOffset(other.m_hOffset)
    , m_vOffset(other.m_vOffset)
    , m_backgroundColor(other.m_backgroundColor)
    , m_showShadow(other.m_showShadow)
    , m_shadowColor(other.m_shadowColor)
    , m_commentSize(other.m_commentSize)
    , m_textAlign(other.m_textAlign)
{

}
CommentFormatPrivate& CommentFormatPrivate::operator=(const CommentFormatPrivate& other)
{
    m_hOffset = other.m_hOffset;
    m_vOffset = other.m_vOffset;
    m_backgroundColor = other.m_backgroundColor;
    m_showShadow = other.m_showShadow;
    m_shadowColor = other.m_shadowColor;
    m_commentSize = other.m_commentSize;
    m_textAlign = other.m_textAlign;
    return *this;
}
CommentFormatPrivate::~CommentFormatPrivate()
{

}

const QSizeF& CommentFormat::commentSize() const
{
    Q_D(const CommentFormat);
    return d->m_commentSize;
}

void CommentFormat::setCommentSize(const QSizeF& val)
{
    Q_D(CommentFormat);
    d->m_commentSize = val;
}

const double& CommentFormat::vOffset() const
{
    Q_D(const CommentFormat);
    return d->m_vOffset;
}

const double& CommentFormat::hOffset() const
{
    Q_D(const CommentFormat);
    return d->m_hOffset;
}

void CommentFormat::setVOffset(const double& val)
{
    Q_D(CommentFormat);
    d->m_vOffset = val;
}

void CommentFormat::setHOffset(const double& val)
{
    Q_D(CommentFormat);
    d->m_hOffset = val;
}

const QColor& CommentFormat::shadowColor() const
{
    Q_D(const CommentFormat);
    return d->m_shadowColor;
}

void CommentFormat::shadowColor(const QColor& val)
{
    Q_D(CommentFormat);
    d->m_shadowColor = val;
}

const bool& CommentFormat::showShadow() const
{
    Q_D(const CommentFormat);
    return d->m_showShadow;
}

void CommentFormat::setShowShadow(const bool& val)
{
    Q_D(CommentFormat);
    d->m_showShadow = val;
}

const QColor& CommentFormat::backgroundColor() const
{
    Q_D(const CommentFormat);
    return d->m_backgroundColor;
}

void CommentFormat::setBackgroundColor(const QColor& val)
{
    Q_D(CommentFormat);
    qDebug() << "Old Col: " + d->m_backgroundColor.name() + " New Col: " + val.name();
    d->m_backgroundColor = val;
    qDebug() << d->m_backgroundColor.name();
}

void CommentFormat::setOffset(const double& h, const double& v)
{
    setHOffset(h);
    setVOffset(v);
}

CommentFormat::CommentFormat()
    :d_ptr(new CommentFormatPrivate(this))
{

}

CommentFormat::CommentFormat(const CommentFormat& other)
    : d_ptr(new CommentFormatPrivate(this))
{
    Q_D(CommentFormat);
    d->operator=(*(other.d_ptr));
}

CommentFormat::~CommentFormat()
{
    if (d_ptr)
        delete d_ptr;
    d_ptr = NULL;
}

CommentFormat& CommentFormat::operator=(const CommentFormat& other)
{
    Q_D(CommentFormat);
    d->operator=(*(other.d_ptr));
    return *this;
}

const Qt::Alignment& CommentFormat::textAlign() const
{
    Q_D(const CommentFormat);
    return d->m_textAlign;
}

void CommentFormat::setTextAlign(const Qt::Alignment& val)
{
    Q_D(CommentFormat);
    d->m_textAlign = val;
}

QT_END_NAMESPACE_XLSX