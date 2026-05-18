#pragma once
#include <QMainWindow>
#include <QLineEdit>
#include <QLabel>
#include <QPushButton>
#include <QPlainTextEdit>
#include <QProcess>
#include <QThread>
#include <QNetworkAccessManager>
#include "logic_handler.h"

class MainWindow : public QMainWindow {
    Q_OBJECT
public:
    explicit MainWindow(QWidget* parent = nullptr);
    ~MainWindow();

private slots:
    void onStart();
    void onStop();
    void onBrowseChrome();
    void onCheckVersions();
    void appendLog(const QString& msg);
    void onTaskFinished(int successCount);
    void checkAndPromptDriverUpdate();
    void startDriverDownload(int majorVersion);

private:
    void initUI();
    void autoDetectPaths();

    // Config inputs
    QLineEdit* m_chromePath = nullptr;
    QLineEdit* m_url = nullptr;
    QLineEdit* m_count = nullptr;
    QLineEdit* m_minWait = nullptr;
    QLineEdit* m_maxWait = nullptr;

    // Version labels
    QLabel* m_chromeVer = nullptr;
    QLabel* m_driverVer = nullptr;

    // Buttons
    QPushButton* m_startBtn = nullptr;
    QPushButton* m_stopBtn = nullptr;

    // Log
    QPlainTextEdit* m_logView = nullptr;

    // Worker
    QThread* m_workerThread = nullptr;
    LogicHandler* m_logicHandler = nullptr;

    // App directory
    QString m_appDir;

    // Network
    QNetworkAccessManager* m_netManager = nullptr;
};
