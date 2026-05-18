#include "main_window.h"
#include "logic_handler.h"

#include <QApplication>
#include <QCoreApplication>
#include <QDir>
#include <QFile>
#include <QFileDialog>
#include <QFileInfo>
#include <QGroupBox>
#include <QHBoxLayout>
#include <QMessageBox>
#include <QProcess>
#include <QTextStream>
#include <QTime>
#include <QVBoxLayout>
#include <QFrame>
#include <QRegularExpression>
#include <QNetworkReply>
#include <QJsonDocument>
#include <QJsonObject>
#include <QJsonArray>
#include <QProgressDialog>
#include <QTimer>
#include <QDirIterator>

MainWindow::MainWindow(QWidget* parent)
    : QMainWindow(parent)
{
    setWindowTitle("CRM 自动化助手v3.0 bySTL  [C++]");
    resize(620, 700);

    m_appDir = QCoreApplication::applicationDirPath();

    m_netManager = new QNetworkAccessManager(this);

    initUI();
    autoDetectPaths();

    QTimer::singleShot(500, this, &MainWindow::checkAndPromptDriverUpdate);
}

MainWindow::~MainWindow()
{
    if (m_logicHandler)
        m_logicHandler->doStop();

    if (m_workerThread) {
        m_workerThread->quit();
        m_workerThread->wait(5000);
    }
}

// ─── UI Construction ─────────────────────────────────────────────

void MainWindow::initUI()
{
    auto* central = new QWidget(this);
    setCentralWidget(central);
    auto* mainLayout = new QVBoxLayout(central);

    // ── Config group ──────────────────────────────────────────
    auto* configGroup = new QGroupBox("运行配置", this);
    auto* configLayout = new QGridLayout(configGroup);

    configLayout->addWidget(new QLabel("Chrome路径:"), 0, 0);
    m_chromePath = new QLineEdit;
    configLayout->addWidget(m_chromePath, 0, 1);
    auto* browseBtn = new QPushButton("...");
    browseBtn->setMaximumWidth(40);
    configLayout->addWidget(browseBtn, 0, 2);
    connect(browseBtn, &QPushButton::clicked, this, &MainWindow::onBrowseChrome);

    configLayout->addWidget(new QLabel("登录网址:"), 1, 0);
    m_url = new QLineEdit(
        "https://www.fxiaoke.com/proj/page/login?returnUrl=https%3A%2F%2Fwww.fxiaoke.com%2FXV%2FUI%2FHome#paasapp/list/=/appId_CRM/AccountObj");
    configLayout->addWidget(m_url, 1, 1, 1, 2);

    configLayout->addWidget(new QLabel("录入数量:"), 2, 0);
    m_count = new QLineEdit("50");
    m_count->setMaximumWidth(80);
    configLayout->addWidget(m_count, 2, 1);

    configLayout->addWidget(new QLabel("随机等待(秒):"), 3, 0);
    auto* timeLayout = new QHBoxLayout;
    m_minWait = new QLineEdit("10");
    m_minWait->setMaximumWidth(50);
    m_maxWait = new QLineEdit("20");
    m_maxWait->setMaximumWidth(50);
    timeLayout->addWidget(m_minWait);
    timeLayout->addWidget(new QLabel("-"));
    timeLayout->addWidget(m_maxWait);
    timeLayout->addWidget(new QLabel("(范围)"));
    timeLayout->addStretch();
    configLayout->addLayout(timeLayout, 3, 1, 1, 2);

    mainLayout->addWidget(configGroup);

    // ── Version info bar ──────────────────────────────────────
    auto* infoFrame = new QFrame(this);
    auto* infoLayout = new QHBoxLayout(infoFrame);
    infoLayout->setContentsMargins(0, 2, 0, 2);

    infoLayout->addWidget(new QLabel("Chrome: "));
    m_chromeVer = new QLabel("检测中...");
    m_chromeVer->setStyleSheet("color: blue;");
    infoLayout->addWidget(m_chromeVer);

    infoLayout->addWidget(new QLabel(" | Driver: "));
    m_driverVer = new QLabel("检测中...");
    m_driverVer->setStyleSheet("color: green;");
    infoLayout->addWidget(m_driverVer);

    infoLayout->addStretch();
    auto* refreshBtn = new QPushButton("刷新版本");
    connect(refreshBtn, &QPushButton::clicked, this, &MainWindow::onCheckVersions);
    infoLayout->addWidget(refreshBtn);

    mainLayout->addWidget(infoFrame);

    // ── Action buttons ────────────────────────────────────────
    auto* actionLayout = new QHBoxLayout;
    m_startBtn = new QPushButton("开始运行");
    m_stopBtn = new QPushButton("停止");
    m_stopBtn->setEnabled(false);

    m_startBtn->setMinimumHeight(36);
    m_stopBtn->setMinimumHeight(36);

    connect(m_startBtn, &QPushButton::clicked, this, &MainWindow::onStart);
    connect(m_stopBtn, &QPushButton::clicked, this, &MainWindow::onStop);

    actionLayout->addWidget(m_startBtn);
    actionLayout->addWidget(m_stopBtn);
    mainLayout->addLayout(actionLayout);

    // ── Log area ──────────────────────────────────────────────
    auto* logGroup = new QGroupBox("日志", this);
    auto* logLayout = new QVBoxLayout(logGroup);
    m_logView = new QPlainTextEdit;
    m_logView->setReadOnly(true);
    m_logView->setMinimumHeight(200);
    logLayout->addWidget(m_logView);
    mainLayout->addWidget(logGroup, 1);
}

// ─── Path detection ──────────────────────────────────────────────

void MainWindow::autoDetectPaths()
{
    QStringList knownPaths = {
        "C:/Program Files/Google/Chrome/Application/chrome.exe",
        "C:/Program Files (x86)/Google/Chrome/Application/chrome.exe",
        QDir::homePath() + "/AppData/Local/Google/Chrome/Application/chrome.exe"
    };

    // Try saved path first
    QString savedPath = m_appDir + "/chrome_path.txt";
    QFile savedFile(savedPath);
    if (savedFile.open(QIODevice::ReadOnly | QIODevice::Text)) {
        QString p = QTextStream(&savedFile).readAll().trimmed().remove('"');
        savedFile.close();
        if (QFileInfo::exists(p)) {
            m_chromePath->setText(p);
            onCheckVersions();
            return;
        }
    }

    for (const auto& p : knownPaths) {
        if (QFileInfo::exists(p)) {
            m_chromePath->setText(p);
            break;
        }
    }
    onCheckVersions();
}

// ─── Version detection ───────────────────────────────────────────

void MainWindow::onCheckVersions()
{
    QString chromePath = m_chromePath->text();
    QString driverPath = m_appDir + "/chromedriver.exe";

    // Chrome version via wmic (Win7 compatible)
    QString cVer = "未检测到";
    if (!chromePath.isEmpty() && QFileInfo::exists(chromePath)) {
        QString escaped = QDir::toNativeSeparators(chromePath);
        escaped.replace("\\", "\\\\");
        QProcess proc;
        proc.start("wmic", {"datafile", "where", "name=\"" + escaped + "\"",
                            "get", "Version", "/value"});
        proc.waitForFinished(5000);
        QString out = QString::fromLocal8Bit(proc.readAllStandardOutput());
        int idx = out.indexOf("Version=");
        if (idx >= 0) {
            cVer = out.mid(idx + 8);
            cVer = cVer.section(QRegularExpression("\\s+"), 0, 0);
        } else if (!out.isEmpty()) {
            cVer = "存在 (读取失败)";
        }
    }
    m_chromeVer->setText(cVer);

    // Driver version
    QString dVer = "未检测到";
    if (QFileInfo::exists(driverPath)) {
        QProcess proc;
        proc.start(driverPath, {"--version"});
        proc.waitForFinished(5000);
        QString out = QString::fromLocal8Bit(proc.readAllStandardOutput());
        QStringList parts = out.split(QRegularExpression("\\s+"));
        if (parts.size() >= 2) {
            dVer = parts[1];
        } else if (!out.isEmpty()) {
            dVer = "存在 (读取失败)";
        }
    }
    m_driverVer->setText(dVer);
}

// ─── File browser ────────────────────────────────────────────────

void MainWindow::onBrowseChrome()
{
    QString path = QFileDialog::getOpenFileName(this, "选择 Chrome 可执行文件",
                                                 QString(), "Exe (*.exe)");
    if (!path.isEmpty()) {
        m_chromePath->setText(path);
        onCheckVersions();

        // Save path
        QFile f(m_appDir + "/chrome_path.txt");
        if (f.open(QIODevice::WriteOnly | QIODevice::Text)) {
            QTextStream(&f) << path;
            f.close();
        }
    }
}

// ─── Logging ─────────────────────────────────────────────────────

void MainWindow::appendLog(const QString& msg)
{
    QString ts = QTime::currentTime().toString("HH:mm:ss");
    m_logView->appendPlainText(QString("[%1] %2").arg(ts, msg));
    auto cursor = m_logView->textCursor();
    cursor.movePosition(QTextCursor::End);
    m_logView->setTextCursor(cursor);
}

// ─── Task control ────────────────────────────────────────────────

void MainWindow::onStart()
{
    if (m_logicHandler) {
        QMessageBox::warning(this, "提示", "任务已在运行中");
        return;
    }

    if (m_chromePath->text().isEmpty()) {
        QMessageBox::critical(this, "错误", "请指定 Chrome 路径");
        return;
    }

    bool ok;
    int limit = m_count->text().toInt(&ok);
    if (!ok) limit = 50;

    int tmin = m_minWait->text().toInt(&ok);
    if (!ok || tmin < 0) {
        QMessageBox::critical(this, "错误", "随机等待时间必须是大于0的数字");
        return;
    }

    int tmax = m_maxWait->text().toInt(&ok);
    if (!ok || tmax < 0) {
        QMessageBox::critical(this, "错误", "随机等待时间必须是大于0的数字");
        return;
    }

    if (tmin > tmax) {
        QMessageBox::critical(this, "错误", "最小等待时间不能大于最大等待时间");
        return;
    }

    RunParams params;
    params.chromePath = m_chromePath->text();
    params.driverPath = m_appDir + "/chromedriver.exe";
    params.url = m_url->text();
    params.limit = limit;
    params.minWait = tmin;
    params.maxWait = tmax;
    params.appDir = m_appDir;

    // Create worker thread and logic handler
    m_workerThread = new QThread(this);
    m_logicHandler = new LogicHandler();
    m_logicHandler->moveToThread(m_workerThread);

    connect(m_logicHandler, &LogicHandler::log, this, &MainWindow::appendLog);
    connect(m_logicHandler, &LogicHandler::finished, this, &MainWindow::onTaskFinished);

    connect(m_logicHandler, &LogicHandler::needConfirm, this,
        [this](const QString& title, const QString& msg, bool* result) {
            *result = (QMessageBox::question(this, title, msg,
                        QMessageBox::Ok | QMessageBox::Cancel) == QMessageBox::Ok);
        }, Qt::BlockingQueuedConnection);

    connect(m_logicHandler, &LogicHandler::needYesNo, this,
        [this](const QString& title, const QString& msg, bool* result) {
            *result = (QMessageBox::question(this, title, msg,
                        QMessageBox::Yes | QMessageBox::No) == QMessageBox::Yes);
        }, Qt::BlockingQueuedConnection);

    connect(m_workerThread, &QThread::started, m_logicHandler, [this, params]() {
        m_logicHandler->doRun(params);
    });

    m_startBtn->setEnabled(false);
    m_stopBtn->setEnabled(true);

    m_workerThread->start();
}

void MainWindow::onStop()
{
    if (m_logicHandler)
        m_logicHandler->doStop();
}

void MainWindow::onTaskFinished(int successCount)
{
    m_startBtn->setEnabled(true);
    m_stopBtn->setEnabled(false);

    if (m_workerThread) {
        m_workerThread->quit();
        m_workerThread->wait(5000);
        m_logicHandler->deleteLater();
        m_workerThread->deleteLater();
        m_logicHandler = nullptr;
        m_workerThread = nullptr;
    }

    if (successCount > 0) {
        QMessageBox::information(this, "完成",
                                 QString("任务结束，共录入 %1 条。").arg(successCount));
    } else if (successCount == 0) {
        appendLog("任务结束，未录入任何记录");
    }
}

void MainWindow::checkAndPromptDriverUpdate()
{
    QString cVerStr = m_chromeVer->text();
    QString dVerStr = m_driverVer->text();
    
    int cMajor = cVerStr.section('.', 0, 0).toInt();
    int dMajor = dVerStr.section('.', 0, 0).toInt();
    
    if (cMajor == 0) return; // Cannot determine chrome version
    if (cMajor < 115) {
        if (dMajor != cMajor) {
            appendLog(QString("Chrome版本 %1 过低，不支持自动更新 ChromeDriver，请手动更新。").arg(cMajor));
        }
        return;
    }
    
    if (cMajor != dMajor) {
        QString msg = QString("发现 Chrome 浏览器大版本为 %1，而本地 ChromeDriver 版本为 %2，版本不匹配会导致程序无法运行。\n\n是否立即自动下载并更新 ChromeDriver？")
                      .arg(cMajor).arg(dMajor == 0 ? "未检测到" : QString::number(dMajor));
        if (QMessageBox::question(this, "驱动更新提示", msg) == QMessageBox::Yes) {
            startDriverDownload(cMajor);
        }
    }
}

void MainWindow::startDriverDownload(int majorVersion)
{
    QString apiUrl = "https://googlechromelabs.github.io/chrome-for-testing/latest-versions-per-milestone-with-downloads.json";
    
    auto* progress = new QProgressDialog("正在获取下载地址...", "取消", 0, 100, this);
    progress->setWindowTitle("更新 ChromeDriver");
    progress->setWindowModality(Qt::WindowModal);
    progress->show();
    
    QUrl url(apiUrl);
    QNetworkRequest req(url);
    QNetworkReply* reply = m_netManager->get(req);
    
    connect(reply, &QNetworkReply::finished, this, [this, reply, progress, majorVersion]() {
        reply->deleteLater();
        if (reply->error() != QNetworkReply::NoError) {
            QMessageBox::critical(this, "错误", "获取版本信息失败：" + reply->errorString());
            progress->close();
            progress->deleteLater();
            return;
        }
        
        QByteArray data = reply->readAll();
        QJsonDocument doc = QJsonDocument::fromJson(data);
        QJsonObject root = doc.object();
        QJsonObject milestones = root.value("milestones").toObject();
        QJsonObject currentMilestone = milestones.value(QString::number(majorVersion)).toObject();
        QJsonObject downloads = currentMilestone.value("downloads").toObject();
        QJsonArray chromedrivers = downloads.value("chromedriver").toArray();
        
        QString downloadUrl;
        for (const QJsonValue& val : chromedrivers) {
            QJsonObject obj = val.toObject();
            if (obj.value("platform").toString() == "win64") {
                downloadUrl = obj.value("url").toString();
                break;
            } else if (obj.value("platform").toString() == "win32" && downloadUrl.isEmpty()) {
                downloadUrl = obj.value("url").toString();
            }
        }
        
        if (downloadUrl.isEmpty()) {
            QMessageBox::critical(this, "错误", "未找到对应的 win64 或 win32 ChromeDriver 下载链接。");
            progress->close();
            progress->deleteLater();
            return;
        }
        
        progress->setLabelText("正在下载 ChromeDriver...");
        
        QUrl dlUrlObj(downloadUrl);
        QNetworkRequest dlReq(dlUrlObj);
        QNetworkReply* dlReply = m_netManager->get(dlReq);
        
        connect(progress, &QProgressDialog::canceled, dlReply, &QNetworkReply::abort);
        
        connect(dlReply, &QNetworkReply::downloadProgress, progress, [progress](qint64 bytesReceived, qint64 bytesTotal) {
            if (bytesTotal > 0) {
                progress->setMaximum(bytesTotal);
                progress->setValue(bytesReceived);
            }
        });
        
        connect(dlReply, &QNetworkReply::finished, this, [this, dlReply, progress]() {
            dlReply->deleteLater();
            progress->close();
            progress->deleteLater();
            
            if (dlReply->error() != QNetworkReply::NoError) {
                if (dlReply->error() != QNetworkReply::OperationCanceledError) {
                    QMessageBox::critical(this, "错误", "下载失败：" + dlReply->errorString());
                }
                return;
            }
            
            QFile tempZip(m_appDir + "/temp_chromedriver.zip");
            if (!tempZip.open(QIODevice::WriteOnly)) {
                QMessageBox::critical(this, "错误", "无法保存下载的文件。");
                return;
            }
            tempZip.write(dlReply->readAll());
            tempZip.close();
            
            appendLog("下载完成，开始解压...");
            
            QString destDir = m_appDir + "/temp_extracted";
            QDir().mkpath(destDir);
            
            QProcess proc;
            proc.start("powershell", {"-Command", QString("Expand-Archive -Force -Path '%1' -DestinationPath '%2'").arg(tempZip.fileName()).arg(destDir)});
            proc.waitForFinished(30000);
            
            QDir d(destDir);
            QString exePath;
            QDirIterator it(destDir, QStringList() << "chromedriver.exe", QDir::Files, QDirIterator::Subdirectories);
            if (it.hasNext()) {
                exePath = it.next();
            }
            
            if (!exePath.isEmpty()) {
                QString targetPath = m_appDir + "/chromedriver.exe";
                if (QFile::exists(targetPath)) QFile::remove(targetPath);
                QFile::copy(exePath, targetPath);
                appendLog("ChromeDriver 更新成功！");
                QMessageBox::information(this, "成功", "ChromeDriver 更新成功！");
            } else {
                appendLog("解压失败或未找到 chromedriver.exe");
                QMessageBox::critical(this, "错误", "解压失败或未找到 chromedriver.exe。");
            }
            
            QFile::remove(tempZip.fileName());
            d.removeRecursively();
            
            onCheckVersions();
        });
    });
}


