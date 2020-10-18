#ifndef CREATE_CHECK_REPORT
#define CREATE_CHECK_REPORT

#include <QtWidgets/QWidget>
#include "ui_CreateCheckReport.h"
#include <QObject>
#include <QAxWidget>
#include <QAxObject>
#include "qcheckbox.h"
#include <cJSON.h>

namespace Ui
{
	class CreateCheckReport;
}
class CreateCheckReport : public QWidget
{
    Q_OBJECT

public:
    CreateCheckReport(QWidget *parent = Q_NULLPTR);
	~CreateCheckReport();

private:
	void onChosePath();

	void onSavePath();

	void initTable();

	void allCheckBox();

	void onCreateReport();

	void onExit();

	void writeWord();


public:
	bool getCheckMessage(QList<QString>& itemList);

	bool fromPersistentSetting(QList<QString>& itemList);

	bool getCurrentMessage(QList<QString>& itemList, QString path);

	bool getPicturePath(QString& path, QString &picturePath);

	bool getJsonDoubleArray(cJSON* jsonObj, const QString & fieldName, double * result);

private:

	QString filePath;

	QString fileResultPath;

	Ui::CreateCheckReport *ui = nullptr;

	bool m_bOpened;
	QString m_strFilePath;
	QAxObject *m_wordDocuments;
	QAxObject *m_wordWidget = nullptr;

	QCheckBox *box = nullptr;
};
#endif // !CREATE_CHECK_REPORT
