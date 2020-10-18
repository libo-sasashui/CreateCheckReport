#include "wordengine.h"
#include "WordEngine.h"
#include "qt_windows.h"

WordEngine::WordEngine(QObject *parent) : QObject(parent)
{
	m_pWord = NULL;
	m_pWorkDocuments = NULL;
	m_pWorkDocument = NULL;

	m_bIsOpen = false;
	m_bNewFile = false;

	HRESULT result = OleInitialize(0);

	if (result != S_OK && result != S_FALSE)
	{
		qDebug() << QString("Could not initialize OLE (error %x)").arg((unsigned int)result);
	}
}

WordEngine::~WordEngine()
{
	//if(m_bIsOpen)
	//    close();

	OleUninitialize();
}

bool WordEngine::Open(QString sFile, bool bVisible)
{
	//新建一个word应用程序
	m_pWord = new QAxObject();
	bool bFlag = m_pWord->setControl("word.Application");
	if (!bFlag)
	{
		return false;
	}
	m_pWord->setProperty("Visible", bVisible);
	//获取所有的工作文档
	QAxObject *document = m_pWord->querySubObject("Documents");
	if (!document)
	{
		return false;
	}
	//以文件template.dot为模版新建一个文档
	document->dynamicCall("Add(QString)", sFile);
	//获取当前激活的文档
	m_pWorkDocument = m_pWord->querySubObject("ActiveDocument");
	if (m_pWorkDocument)
		m_bIsOpen = true;
	else
		m_bIsOpen = false;

	return m_bIsOpen;
}

void WordEngine::save(QString sSavePath)
{
	if (m_bIsOpen && m_pWorkDocument)
	{
		if (m_bNewFile) {
			m_pWorkDocument->dynamicCall("Save()");
		}
		else {
			//m_pWorkDocument->dynamicCall("SaveAs (const QString&,int,const QString&,const QString&,bool,bool)",
			 //                           m_sFile,56,QString(""),QString(""),false,false);
			m_pWorkDocument->dynamicCall("SaveAs (const QString&)", sSavePath);
		}
	}
	qDebug() << "save Done.";
}

void WordEngine::close(bool bSave)
{
	if (bSave) {
		//save();
	}
	if (m_pWord) {
		m_pWord->setProperty("DisplayAlerts", true);
	}
	if (m_pWorkDocument) {
		m_pWorkDocument->dynamicCall("Close(bool)", true);
	}
	if (m_pWord) {
		m_pWord->dynamicCall("Quit()");
	}
	if (m_pWorkDocuments)
	{
		delete m_pWorkDocuments;
	}
	if (m_pWord)
	{
		delete m_pWord;
	}
	m_pWorkDocument = NULL;
	m_pWorkDocuments = NULL;
	m_pWord = NULL;

	m_bIsOpen = false;
	m_bNewFile = false;

}

bool WordEngine::replaceText(QString sLabel, QString sText)
{
	if (!m_pWorkDocument) {
		return false;
	}
	//获取文档中名字为sLabel的标签
	QAxObject *pBookmark = m_pWorkDocument->querySubObject("Bookmarks(QString)", sLabel);
	if (pBookmark)
	{
		pBookmark->dynamicCall("Select(void)");
		pBookmark->querySubObject("Range")->setProperty("Text", sText);
		delete pBookmark;
	}
	return true;
}

bool WordEngine::replacePic(QString sLabel, QString sFile)
{
	if (!m_pWorkDocument)
		return false;

	QAxObject *bookmark_pic = m_pWorkDocument->querySubObject("Bookmarks(QString)", sLabel);
	if (bookmark_pic)
	{
		bookmark_pic->dynamicCall("Select(void)");
		QAxObject *Inlineshapes = m_pWorkDocument->querySubObject("InlineShapes");
		Inlineshapes->dynamicCall("AddPicture(const QString&)", sFile);
		delete Inlineshapes;
	}
	return true;
}

//QAxObject *WordEngine::insertTable(QString sLabel, int row, int column)
//{
//     QAxObject *bookmark = m_pWorkDocument->querySubObject("Bookmarks(QVariant)", sLabel);
//     if(bookmark)
//     {
//       bookmark->dynamicCall("Select(void)");
//       QAxObject *selection = m_pWord->querySubObject("Selection");

//       selection->dynamicCall("InsertAfter(QString&)", "\n");
//       //selection->dynamicCall("MoveLeft(int)", 1);
//       selection->querySubObject("ParagraphFormat")->dynamicCall("Alignment", "wdAlignParagraphCenter");
//       //selection->dynamicCall("TypeText(QString&)", "Table Test");//设置标题

//       QAxObject *range = selection->querySubObject("Range");
//       QAxObject *tables = m_pWorkDocument->querySubObject("Tables");
//       QAxObject *table = tables->querySubObject("Add(QVariant,int,int)",range->asVariant(),row,column);

//       for(int i=1;i<=6;i++)
//       {
//           QString str = QString("Borders(-%1)").arg(i);

//           QAxObject *borders = table->querySubObject(str.toLocal8Bit().constData());
//           borders->dynamicCall("SetLineStyle(int)",1);
//       }
//       return table;
//     }
//     return NULL;
//}

QAxObject *WordEngine::insertTable(QString sLabel, int row, int column, QStringList headList)
{
	QAxObject *bookmark = m_pWorkDocument->querySubObject("Bookmarks(QVariant)", sLabel);
	if (headList.size() != column) {
		return NULL;
	}
	if (bookmark)
	{
		bookmark->dynamicCall("Select(void)");
		QAxObject *selection = m_pWord->querySubObject("Selection");

		selection->dynamicCall("InsertAfter(QString&)", "\r\n");
		//selection->dynamicCall("MoveLeft(int)", 1);
		selection->querySubObject("ParagraphFormat")->dynamicCall("Alignment", "wdAlignParagraphCenter");
		//设置标题
		//selection->dynamicCall("TypeText(QString&)", "Table Test");

		QAxObject *range = selection->querySubObject("Range");
		QAxObject *tables = m_pWorkDocument->querySubObject("Tables");
		QAxObject *table = tables->querySubObject("Add(QVariant,int,int)", range->asVariant(), row, column);
		//表格自动拉伸列 0固定  1根据内容调整  2 根据窗口调整
		table->dynamicCall("AutoFitBehavior(WdAutoFitBehavior)", 2);

		//设置表头
		for (int i = 0; i < headList.size(); i++) {
			table->querySubObject("Cell(int,int)", 1, i + 1)->querySubObject("Range")->dynamicCall("SetText(QString)", headList.at(i));
			//加粗
			table->querySubObject("Cell(int,int)", 1, i + 1)->querySubObject("Range")->dynamicCall("SetBold(int)", true);
		}

		for (int i = 1; i <= 6; i++)
		{
			QString str = QString("Borders(-%1)").arg(i);
			QAxObject *borders = table->querySubObject(str.toLocal8Bit().constData());
			borders->dynamicCall("SetLineStyle(int)", 1);
		}
		return table;
	}
	return NULL;
}
/**
 * @brief WordEngine::insertTable
 * @param sLabel 模板中的标签
 * @param column 某行最长列
 * @param headList 数据 里面存放这每一行的数据
 * @return
 */
QAxObject *WordEngine::insertTable(QString sLabel, int column, QList<QStringList> headList)
{
	QAxObject *bookmark = m_pWorkDocument->querySubObject("Bookmarks(QVariant)", sLabel);

	if (bookmark)
	{
		bookmark->dynamicCall("Select(void)");
		QAxObject *selection = m_pWord->querySubObject("Selection");

		selection->dynamicCall("InsertAfter(QString&)", "\r\n");
		//selection->dynamicCall("MoveLeft(int)", 1);
		selection->querySubObject("ParagraphFormat")->dynamicCall("Alignment", "wdAlignParagraphCenter");
		//设置标题
		//selection->dynamicCall("TypeText(QString&)", "Table Test");

		QAxObject *range = selection->querySubObject("Range");
		QAxObject *tables = m_pWorkDocument->querySubObject("Tables");
		QAxObject *table = tables->querySubObject("Add(QVariant,int,int)", range->asVariant(), headList.size(), column);
		//表格自动拉伸列 0固定  1根据内容调整  2 根据窗口调整
		table->dynamicCall("AutoFitBehavior(WdAutoFitBehavior)", 2);

		//设置表头和数据
		for (int i = 0; i < headList.size(); i++) {
			QStringList columnList = headList.at(i);
			for (int j = 0; j < columnList.size(); j++) {
				table->querySubObject("Cell(int,int)", i + 1, j + 1)->querySubObject("Range")->dynamicCall("SetText(QString)", columnList.at(j));
				//加粗
				table->querySubObject("Cell(int,int)", 1, j + 1)->querySubObject("Range")->dynamicCall("SetBold(int)", true);
			}
		}

		for (int i = 1; i <= 6; i++)
		{
			QString str = QString("Borders(-%1)").arg(i);
			QAxObject *borders = table->querySubObject(str.toLocal8Bit().constData());
			borders->dynamicCall("SetLineStyle(int)", 1);
		}
		return table;
	}
	return NULL;
}


void WordEngine::setColumnWidth(QAxObject *table, int column, int width)
{
	if (!table) {
		return;
	}
	table->querySubObject("Columns(int)", column)->setProperty("Width", width);
}

void WordEngine::SetTableCellString(QAxObject *table, int row, int column, QString text)
{
	if (!table)
		return;
	QAxObject *cell = table->querySubObject("Cell(int,int)", row, column);
	if (!cell)
		return;
	cell->dynamicCall("Select(void)");
	cell->querySubObject("Range")->setProperty("Text", text);
}