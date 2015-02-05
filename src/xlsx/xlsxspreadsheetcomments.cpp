#include "xlsxspreadsheetcomments.h"
#include "xlsxspreadsheetcomments_p.h"
#include "xlsxcellreference.h"
#include <QXmlStreamWriter>
QT_BEGIN_NAMESPACE_XLSX
SpreadSheetCommentPrivate::SpreadSheetCommentPrivate(SpreadSheetComment* p, SpreadSheetComment::CreateFlag flag) 
: AbstractOOXmlFilePrivate(p,flag)
{
}
const Comment* SpreadSheetCommentPrivate::getComment(int row, int col) const {
	auto colList = m_Comments.value(row, nullptr);
	if (!colList) return nullptr;
	return colList->value(col, nullptr);
}
SpreadSheetCommentPrivate::~SpreadSheetCommentPrivate() {
	for (auto i = m_Comments.begin(); i != m_Comments.end(); ++i) {
		for (auto j = i.value()->begin(); j != i.value()->end(); ++j) {
			delete j.value();
		}
		delete i.value();
	}
}
void SpreadSheetCommentPrivate::setComment(int row, int col, const Comment& val) {
	auto comRow = m_Comments.find(row);
	if (comRow == m_Comments.end()) {
		comRow = m_Comments.insert(row, new QMap<int, Comment*>);
	}
	auto comCol = comRow.value()->find(col);
	if (comCol == comRow.value()->end()) {
		comRow.value()->insert(col, new Comment(val));
	}
	else {
		comCol.value()->operator=(val);
	}
}
void SpreadSheetCommentPrivate::removeComment(int row, int col) {
	auto comRow = m_Comments.find(row);
	if (comRow == m_Comments.end()) return;
	auto comCol = comRow.value()->find(col);
	if (comCol == comRow.value()->end()) return;
	delete comCol.value();
	comRow.value()->erase(comCol);
	if (comRow.value()->isEmpty()) {
		delete comRow.value();
		m_Comments.erase(comRow);
	}
}
void SpreadSheetCommentPrivate::moveComment(int fromRow, int fromCol, int toRow, int toCol) {
	auto OldComment = getComment(fromRow, fromCol);
	if (!OldComment) return;
	setComment(toRow, toCol, *OldComment);
	removeComment(fromRow, fromCol);
}
bool SpreadSheetCommentPrivate::hasRow(int row) const {
	return m_Comments.contains(row);
}
bool SpreadSheetCommentPrivate::isEmpty() const {
	return m_Comments.isEmpty();
}
auto SpreadSheetCommentPrivate::comments() const -> const decltype(m_Comments)&{
	return m_Comments;
}
Comment SpreadSheetComment::getComment(int row, int col) const {
	Q_D(const SpreadSheetComment);
	auto res= d->getComment(row, col);
	if (res) return *res;
	return Comment();
}

void SpreadSheetComment::setComment(int row, int col, const Comment& val) {
	Q_D(SpreadSheetComment);
	d->setComment(row, col, val);
}

void SpreadSheetComment::setComment(const CellReference& rowCol, const Comment& val) {
	if (!rowCol.isValid()) return;
	setComment(rowCol.row(), rowCol.column(), val);
}

void SpreadSheetComment::removeComment(int row, int col) {
	Q_D(SpreadSheetComment);
	d->removeComment(row, col);
}

void SpreadSheetComment::removeComment(const CellReference& rowCol) {
	if (!rowCol.isValid()) return;
	removeComment(rowCol.row(), rowCol.column());
}

void SpreadSheetComment::moveComment(int fromRow, int fromCol, int toRow, int toCol) {
	Q_D(SpreadSheetComment);
	d->moveComment(fromRow, fromCol, toRow, toCol);
}

void SpreadSheetComment::moveComment(const CellReference& fromRowCol, int toRow, int toCol) {
	if (!fromRowCol.isValid()) return;
	moveComment(fromRowCol.row(), fromRowCol.column(), toRow, toCol);
}

void SpreadSheetComment::moveComment(const CellReference& fromRowCol, const CellReference& toRowCol) {
	if (!fromRowCol.isValid() || !toRowCol.isValid()) return;
	moveComment(fromRowCol.row(), fromRowCol.column(), toRowCol.row(), toRowCol.column());
}

void SpreadSheetComment::moveComment(int fromRow, int fromCol, const CellReference& toRowCol) {
	if (!toRowCol.isValid()) return;
	moveComment(fromRow, fromCol, toRowCol.row(), toRowCol.column());
}

void SpreadSheetComment::saveToXmlFile(QIODevice *device) const {
	Q_D(const SpreadSheetComment);
	if (d->isEmpty()) return;
	QXmlStreamWriter writer(device);
	writer.writeStartDocument(QStringLiteral("1.0"), true);
	writer.writeStartElement(QStringLiteral("comments"));
	writer.writeAttribute(QStringLiteral("xmlns"), QStringLiteral("http://schemas.openxmlformats.org/spreadsheetml/2006/main"));
	writer.writeStartElement(QStringLiteral("authors"));
	QStringList authorsList;
	for (auto comRowIter = d->comments().constBegin(); comRowIter != d->comments().constEnd(); ++comRowIter) {
		for (auto comColIter = comRowIter.value()->constBegin(); comColIter != comRowIter.value()->constEnd(); ++comColIter) {
			if (!authorsList.contains(comColIter.value()->author(), Qt::CaseInsensitive))
				authorsList.append(comColIter.value()->author());
		}
	}

	for (auto authIter = authorsList.constBegin(); authIter != authorsList.constEnd(); ++authIter) {
		writer.writeStartElement(QStringLiteral("author"));
		writer.writeCharacters(*authIter);
		writer.writeEndElement();//author
	}
	writer.writeEndElement();//authors
	writer.writeStartElement(QStringLiteral("commentList"));
	QStringList commentLines;
	for (auto comRowIter = d->comments().constBegin(); comRowIter != d->comments().constEnd(); ++comRowIter) {
		for (auto comColIter = comRowIter.value()->constBegin(); comColIter != comRowIter.value()->constEnd(); ++comColIter) {
			writer.writeStartElement(QStringLiteral("comment"));
			writer.writeAttribute(QStringLiteral("ref"), CellReference(comRowIter.key(), comColIter.key()).toString());
			writer.writeAttribute(QStringLiteral("authorId"), QString::number(authorsList.indexOf(QRegExp(comColIter.value()->author(), Qt::CaseInsensitive))));
			writer.writeStartElement(QStringLiteral("text"));
			commentLines = comColIter.value()->text().toPlainString().split(QRegExp("(?:\r\n)|\r|\n"));
			for (auto comLinIter = commentLines.constBegin(); comLinIter != commentLines.constEnd(); ++comLinIter) {
				writer.writeStartElement(QStringLiteral("r"));
				//////////////////////////////////////////////////////////////////////////
				writer.writeStartElement(QStringLiteral("rPr"));
				writer.writeEmptyElement(QStringLiteral("sz"));
				writer.writeAttribute(QStringLiteral("val"), QStringLiteral("9"));
				writer.writeEmptyElement(QStringLiteral("color"));
				writer.writeAttribute(QStringLiteral("indexed"), QStringLiteral("81"));
				writer.writeEmptyElement(QStringLiteral("rFont"));
				writer.writeAttribute(QStringLiteral("val"), QStringLiteral("Tahoma"));
				writer.writeEmptyElement(QStringLiteral("family"));
				writer.writeAttribute(QStringLiteral("val"), QStringLiteral("2"));
				writer.writeEndElement();//rPr
				//////////////////////////////////////////////////////////////////////////
				writer.writeStartElement(QStringLiteral("t"));
				writer.writeAttribute(QStringLiteral("xml:space"), QStringLiteral("preserve"));
				writer.writeCharacters(*comLinIter);
				writer.writeEndElement();//t
				writer.writeEndElement();//r
			}
			writer.writeEndElement();//text
			writer.writeEndElement();//comment
		}
	}
	writer.writeEndElement();//commentList
	writer.writeEndElement();//comments
}

bool SpreadSheetComment::loadFromXmlFile(QIODevice *device) {
	return true;
}
SpreadSheetComment::SpreadSheetComment(CreateFlag flag)
	: AbstractOOXmlFile(new SpreadSheetCommentPrivate(this, flag))
{
}

bool SpreadSheetComment::hasComment(int row, int col) const {
	Q_D(const SpreadSheetComment);
	return d->getComment(row, col);
}

bool SpreadSheetComment::hasComment(const CellReference& rowCol) const {
	if (rowCol.isValid()) return hasComment(rowCol.row(), rowCol.column());
	return false;
}

bool SpreadSheetComment::hasRow(int row) const {
	Q_D(const SpreadSheetComment);
	return d->hasRow(row);
}

bool SpreadSheetComment::isEmpty() const {
	Q_D(const SpreadSheetComment);
	return d->isEmpty();
}

SpreadSheetComment::~SpreadSheetComment() {
}


QT_END_NAMESPACE_XLSX