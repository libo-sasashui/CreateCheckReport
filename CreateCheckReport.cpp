#include "CreateCheckReport.h"
#include "qfiledialog.h"
#include "qfileinfo.h"
#include <cJSON.h>
#include <windows.h> 
#include <fstream>
#include "wordengine.h"
#include <QDir>
#include <QMessageBox>
#include <iostream>
#include <string>

using namespace std;

CreateCheckReport::CreateCheckReport(QWidget *parent)
	:QWidget(parent),
	ui(new Ui::CreateCheckReport)
{
    ui->setupUi(this);

	ui->tableWidget->setWindowFlags(Qt::FramelessWindowHint);

	ui->tableWidget->horizontalHeader()->setSectionResizeMode(0, QHeaderView::ResizeToContents);

	/*设置表格是否充满，即行末不留空*/
	ui->tableWidget->horizontalHeader()->setStretchLastSection(true);

	ui->tableWidget->setSelectionMode(QAbstractItemView::NoSelection);
	// 禁止编辑
	ui->tableWidget->setEditTriggers(QAbstractItemView::NoEditTriggers);

	ui->tableWidget->horizontalHeader()->setVisible(false);

	connect(ui->toolButton_savePath, &QAbstractButton::clicked, this, &CreateCheckReport::onChosePath);
	connect(ui->toolButton_resultPath, &QAbstractButton::clicked, this, &CreateCheckReport::onSavePath);
	connect(ui->pushButton_ok, &QAbstractButton::clicked, this, &CreateCheckReport::onCreateReport);
	connect(ui->pushButton_close, &QAbstractButton::clicked, this, &CreateCheckReport::onExit);
	//connect(ui->)
}

CreateCheckReport::~CreateCheckReport()
{
	delete ui;
	ui = nullptr;
}

void CreateCheckReport::onChosePath()
{
	 filePath = QFileDialog::getExistingDirectory(this, QStringLiteral("选择工程文件"));
	if (filePath.isEmpty())
	{
		return;
	}
	else
	{
		ui->lineEdit_projectPath->setText(filePath);
	}
	initTable();
}

void CreateCheckReport::onSavePath()
{
	fileResultPath = QFileDialog::getExistingDirectory(this, QStringLiteral("选择保存路径"));
	if (fileResultPath.isEmpty())
	{
		return;
	}
	else
	{
		ui->lineEdit_resultPath->setText(fileResultPath);
	}
	
}

void CreateCheckReport::initTable()
{
	int row = ui->tableWidget->rowCount();
	for (int i = 1; i <= row; i++)
	{
		ui->tableWidget->removeRow(0);
	}
	QList<QString> itemList;
	fromPersistentSetting(itemList);
	
	int rowCount = ui->tableWidget->rowCount();

	for (size_t i = 0; i < itemList.size(); i++)
	{
		ui->tableWidget->insertRow(i);
		ui->tableWidget->setItem(i, 1, new QTableWidgetItem(itemList[i]));
	}
	//设置表格首列为CheckBox控件
	for (int i = 0; i < ui->tableWidget->rowCount(); i++)
	{
		QTableWidgetItem* p_check = new QTableWidgetItem();
		p_check->setCheckState(Qt::Unchecked); 
		ui->tableWidget->setItem(i, 0, p_check);
	}

}

void CreateCheckReport::allCheckBox()
{
}

void CreateCheckReport::onExit()
{
	this->close();
}


bool CreateCheckReport::fromPersistentSetting(QList<QString>& itemList)
{
	QString fileName = QStringLiteral("%1/类别管理.json").arg(filePath);
	QFile file(fileName);
	if (file.exists() == false)
	{
		file.open(QIODevice::WriteOnly);
		file.close();
	}

	itemList.clear();

	qint64 len = QFileInfo(fileName).size();
	std::vector<char> idx_input;
	idx_input.resize(len + 1);
	std::ifstream ifs(fileName.toLocal8Bit().constData(), std::ios::in);
	if (!ifs.good())
	{
		return false;
	}
	ifs.read(&(idx_input.front()), len);

	idx_input.back() = '\0';
	ifs.close();

	std::shared_ptr<cJSON> idx(cJSON_Parse(&(idx_input.front())), cJSON_Delete);//cjson_parse返回一个可以查询的cJSON对象
	cJSON* jsonObj = idx.get();
	if (!jsonObj)
	{
		return true;
	}
	int itemnumber = cJSON_GetArraySize(jsonObj);

	for (int i = 0; i < itemnumber; ++i)
	{
		cJSON* jsonArray = cJSON_GetArrayItem(jsonObj, i);
		cJSON *code = cJSON_GetArrayItem(jsonArray, 0);

		QString type = QString::fromLocal8Bit(code->valuestring);
		itemList.append(type);
	}

	return false;
}

void CreateCheckReport::onCreateReport()
{
	if (ui->lineEdit_projectPath->text().isEmpty())
	{
		QMessageBox::warning(this, QStringLiteral("提示"), QStringLiteral("请选择保存路径"));
		return;
	}
	if (ui->lineEdit_resultPath->text().isEmpty())
	{
		QMessageBox::warning(this, QStringLiteral("提示"), QStringLiteral("请选择保存路径"));
		return;
	}

	writeWord();
}
bool CreateCheckReport::getCheckMessage(QList<QString>& itemList)
{
	
	QString fileName = QStringLiteral("%1/质检人信息.json").arg(filePath);
	QFile file(fileName);
	if (file.exists() == false)
	{
		file.open(QIODevice::WriteOnly);
		file.close();
	}

	itemList.clear();

	qint64 len = QFileInfo(fileName).size();
	std::vector<char> idx_input;
	idx_input.resize(len + 1);
	std::ifstream ifs(fileName.toLocal8Bit().constData(), std::ios::in);
	if (!ifs.good())
	{
		return false;
	}
	ifs.read(&(idx_input.front()), len);

	idx_input.back() = '\0';
	ifs.close();

	std::shared_ptr<cJSON> idx(cJSON_Parse(&(idx_input.front())), cJSON_Delete);//cjson_parse返回一个可以查询的cJSON对象
	cJSON* jsonObj = idx.get();
	if (!jsonObj)
	{
		return true;
	}
	int itemnumber = cJSON_GetArraySize(jsonObj);

	for (int i = 0; i < itemnumber; ++i)
	{
		cJSON* jsonArray = cJSON_GetArrayItem(jsonObj, i);
		cJSON *projectName = cJSON_GetArrayItem(jsonArray, 0);
		cJSON *checkMan = cJSON_GetArrayItem(jsonArray, 1);
		cJSON *remark = cJSON_GetArrayItem(jsonArray, 2);

		QString typeProjectName = QString::fromLocal8Bit(projectName->valuestring);
		QString typeCheckMan = QString::fromLocal8Bit(checkMan->valuestring);
		QString typeRemark = QString::fromLocal8Bit(remark->valuestring);
		
		itemList.append(typeProjectName);
		itemList.append(typeCheckMan);
		itemList.append(typeRemark);
	}

	return false;
}


bool CreateCheckReport::getCurrentMessage(QList<QString>& itemList, QString path)
{
	QFile file(path);
	if (file.exists() == false)
	{
		file.open(QIODevice::WriteOnly);
		file.close();
	}

	itemList.clear();

	qint64 len = QFileInfo(path).size();
	std::vector<char> idx_input;
	idx_input.resize(len + 1);
	std::ifstream ifs(path.toLocal8Bit().constData(), std::ios::in);
	if (!ifs.good())
	{
		return false;
	}
	ifs.read(&(idx_input.front()), len);

	idx_input.back() = '\0';
	ifs.close();

	std::shared_ptr<cJSON> idx(cJSON_Parse(&(idx_input.front())), cJSON_Delete);//cjson_parse返回一个可以查询的cJSON对象
	cJSON* jsonObj = idx.get();
	if (!jsonObj)
	{
		return true;
	}
	int itemnumber = cJSON_GetArraySize(jsonObj);


	double point[3] = { 0 };

	for (int i = 0; i < itemnumber; ++i)
	{
		cJSON* jsonArray = cJSON_GetArrayItem(jsonObj, i);
		cJSON *problemType = cJSON_GetArrayItem(jsonArray, 0);
		cJSON *problemRemark = cJSON_GetArrayItem(jsonArray, 1);
	
		getJsonDoubleArray(jsonArray, "point", point);
		
		QString ProblemType = QString::fromLocal8Bit(problemType->valuestring);
		QString ProblemRemark = QString::fromLocal8Bit(problemRemark->valuestring);

		double pointX = point[0];
		double pointY = point[1];
		double pointZ = point[2];

		QString xPoint = QString::number(pointX, 'f',16);
		QString yPoint = QString::number(pointY, 'f',16);
		QString zPoint = QString::number(pointZ, 'f',16);

		itemList.append(ProblemType);
		itemList.append(ProblemRemark);
		itemList.append(xPoint);
		itemList.append(yPoint);
		itemList.append(zPoint);
	}

	return false;

}

bool CreateCheckReport::getJsonDoubleArray(cJSON* jsonObj, const QString & fieldName, double * result)
{
	cJSON* jsonArray = cJSON_GetObjectItem(jsonObj, fieldName.toLocal8Bit().constData());
	if (jsonArray)
	{
		for (int i = 0; i < cJSON_GetArraySize(jsonArray); i++)
		{
			result[i] = cJSON_GetArrayItem(jsonArray, i)->valuedouble;
		}

		return true;
	}
	else
	{
		return false;
	}
}

bool CreateCheckReport::getPicturePath(QString& path,QString &picturePath)
{
	QString fileDirPath = filePath + "/" + path;

	QDir fileDir(fileDirPath);
	QFileInfoList list = fileDir.entryInfoList(QStringList() << "*.jpg");

	picturePath = QStringLiteral("%1/%2/.jpg").arg(fileDirPath).arg(list.at(0).baseName());

	return false;
}


void CreateCheckReport::writeWord()
{
	QString runPath = QCoreApplication::applicationDirPath();
	QString templatePath = QStringLiteral("%1/config/report.doc").arg(runPath);
	WordEngine *basicsMessageWord = new WordEngine();
	basicsMessageWord->Open(templatePath, false);

	QList<QString> itemList;
	getCheckMessage(itemList);
	QString projectName = itemList[0];
	QString checkMan = itemList[1];
	QString remark = itemList[2];
	//存入质检基础信息
	basicsMessageWord->replaceText("projectName", projectName);
	basicsMessageWord->replaceText("checkMan", checkMan);
	basicsMessageWord->replaceText("remark", remark);

	QDir *temp = new QDir;
	QString tempPath = QStringLiteral("%1/%2").arg(fileResultPath).arg(projectName);
	bool exist = temp->exists(tempPath);
	if (!exist)
	{
		temp->mkdir(tempPath);
	}

	QString cheakReportPath = QStringLiteral("%1/CheckReport.doc").arg(tempPath);


	basicsMessageWord->save(cheakReportPath);
	basicsMessageWord->close();

	//获取checkbox上选中的类别
	QList<QString> tableWidgetList;
	for (int i = 0; i < ui->tableWidget->rowCount(); i++)
	{
		if (ui->tableWidget->item(i, 0)->checkState() == Qt::Checked)
		{
			//选中执行的操作
			QString str = ui->tableWidget->item(i, 1)->text();
			tableWidgetList.append(str);
		}
	}

	//获取每个类别对应的文件夹，然后读取其中的数据
	int currentChoice = tableWidgetList.size();

	for (int j = 0; j < currentChoice; j++)
	{
		QList<QString>problemList;
		QString currentChoiceName = tableWidgetList[j];

		QString fileDirPath = QStringLiteral("%1/%2").arg(filePath).arg(currentChoiceName);

		QDir fileDir(fileDirPath);
		QFileInfoList list = fileDir.entryInfoList(QStringList() << "*.para");
		for (int z = 0;z < list.size();z++)
		{
			int num = z + 1;
			WordEngine *checkProblemWord = new WordEngine();
		
			QString checkReportTemplatePath = QStringLiteral("%1/config/checkReport.doc").arg(runPath);

			checkProblemWord->Open(checkReportTemplatePath, false);

			QString fileName = QStringLiteral("%1/%2.para").arg(fileDirPath).arg(list.at(z).baseName());

			getCurrentMessage(problemList, fileName);
			QString problemType = problemList[0];
			QString problemRemark = problemList[1];
			QString pointX = problemList[2];
			QString pointY = problemList[3];
			QString pointZ = problemList[4];

			checkProblemWord->replaceText("type", problemType);
			checkProblemWord->replaceText("typeRemark", problemRemark);
			checkProblemWord->replaceText("x", pointX);
			checkProblemWord->replaceText("y", pointY);
			checkProblemWord->replaceText("z", pointZ);

			QString picturePath = QStringLiteral("%1/%2.jpg").arg(fileDirPath).arg(list.at(z).baseName());
			checkProblemWord->replacePic("picture", picturePath);

			QString cheakProblemReportPath = QStringLiteral("%1/%2%3.doc").arg(tempPath).arg(currentChoiceName).arg(num);
			checkProblemWord->save(cheakProblemReportPath);
			checkProblemWord->close();
		}
	}
	QStringList argu;
	argu.append(tempPath);
	argu.append(projectName+".doc");

	//剥离式启动进程
	QProcess process;
	QString program = "UnionWord.exe";
	QProcess::execute(program, argu);
	QMessageBox::information(this, QStringLiteral("提示"), QStringLiteral("导出完成"));
	
}
