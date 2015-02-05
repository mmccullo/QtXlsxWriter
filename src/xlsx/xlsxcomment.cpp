#include "xlsxcomment.h"
#ifndef QT_NO_DEBUG_STREAM
#include <QDebug>
#endif
QT_BEGIN_NAMESPACE_XLSX
Comment::Comment(const QString& auth, const RichString& txt)
	:m_Author(auth)
	, m_Text(txt) {}
Comment::Comment(const Comment &other)
	: m_Author(other.m_Author)
	, m_Text(other.m_Text)
{}

Comment::Comment() {}

Comment& Comment::operator=(const Comment &other) {
	m_Author = other.m_Author;
	m_Text = other.m_Text;
	return *this;
}

const QString& Comment::author() const {
	return m_Author;
}

const RichString& Comment::text() const {
	return m_Text;
}
void Comment::setAuthor(const QString& auth) {
	m_Author = auth;
}

void Comment::setText(const RichString& txt) {
	m_Text = txt;
}
Comment::~Comment() { }

bool operator==(const Comment &rs1, const Comment &rs2) {
	return
		rs1.m_Author.compare(rs2.m_Author,Qt::CaseInsensitive) == 0
		&& rs1.m_Text == rs2.m_Text
	;
}

bool operator!=(const Comment &rs1, const Comment &rs2) {
	return !(rs1 == rs2);
}
#ifndef QT_NO_DEBUG_STREAM
Q_XLSX_EXPORT QDebug operator<<(QDebug dbg, const Comment &rs) {
	return dbg << "Comment Author: " + rs.author() + " Comment Text: " + rs.text().toHtml();
}
#endif
QT_END_NAMESPACE_XLSX
