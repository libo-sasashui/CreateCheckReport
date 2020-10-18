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

	/// ��Word�ļ������sFile·��Ϊ�ջ��������µ�Word�ĵ�
	bool Open(QString sFile, bool bVisible = true);

	void save(QString sSavePath);

	void close(bool bSave = true);

	bool replaceText(QString sLabel, QString sText);
	bool replacePic(QString sLabel, QString sFile);
	//����һ�����м��б��
//    QAxObject *insertTable(QString sLabel,int row,int column);
	//����һ�����м��б�� �����ñ�ͷ
	QAxObject *insertTable(QString sLabel, int row, int column, QStringList headList);

	QAxObject *insertTable(QString sLabel, int column, QList<QStringList> headList);
	//�����п�
	void setColumnWidth(QAxObject *table, int column, int width);
	void SetTableCellString(QAxObject *table, int row, int column, QString text);
signals:

public slots:

private:
	QAxObject *m_pWord;      //ָ������WordӦ�ó���
	QAxObject *m_pWorkDocuments;  //ָ���ĵ���,Word�кܶ��ĵ�
	QAxObject *m_pWorkDocument;   //ָ��m_sFile��Ӧ���ĵ�������Ҫ�������ĵ�

	QString m_sFile;
	bool m_bIsOpen;
	bool m_bNewFile;
};

#endif // WORDENGINE_H