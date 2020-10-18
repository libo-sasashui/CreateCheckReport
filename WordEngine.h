#ifndef WORDENGINE_H
#define WORDENGINE_H
#include <QObject>
#include <QAxObject>
#include <QtCore>

class WordEngine : public QObject
{
	Q_OBJECT
public:
	explicit WordEngine(QObject *parent = 0);
	//WordEngine();
	~WordEngine();

	/// 打开Word文件，如果sFile路径为空或错误，则打开新的Word文档
	bool Open(QString sFile, bool bVisible = true);

	void save(QString sSavePath);

	void close(bool bSave = true);

	bool replaceText(QString sLabel, QString sText);
	bool replacePic(QString sLabel, QString sFile);
	//插入一个几行几列表格
//    QAxObject *insertTable(QString sLabel,int row,int column);
	//插入一个几行几列表格 并设置表头
	QAxObject *insertTable(QString sLabel, int row, int column, QStringList headList);

	QAxObject *insertTable(QString sLabel, int column, QList<QStringList> headList);
	//设置列宽
	void setColumnWidth(QAxObject *table, int column, int width);
	void SetTableCellString(QAxObject *table, int row, int column, QString text);
signals:

public slots:

private:
	QAxObject *m_pWord;      //指向整个Word应用程序
	QAxObject *m_pWorkDocuments;  //指向文档集,Word有很多文档
	QAxObject *m_pWorkDocument;   //指向m_sFile对应的文档，就是要操作的文档

	QString m_sFile;
	bool m_bIsOpen;
	bool m_bNewFile;
};

#endif // WORDENGINE_H