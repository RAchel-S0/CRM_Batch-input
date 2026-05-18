#pragma once
#include <QObject>
#include <QString>
#include <QStringList>
#include <QAtomicInt>

struct RunParams {
    QString chromePath;
    QString driverPath;
    QString url;
    int limit = 50;
    int minWait = 10;
    int maxWait = 20;
    QString appDir;
};

class LogicHandler : public QObject {
    Q_OBJECT
public:
    explicit LogicHandler(QObject* parent = nullptr);

public slots:
    void doRun(const RunParams& params);
    void doStop();

signals:
    void log(const QString& message);
    void finished(int successCount);
    void needConfirm(const QString& title, const QString& message, bool* result);
    void needYesNo(const QString& title, const QString& message, bool* result);

private:
    QAtomicInt m_running;
    QStringList m_replyPool;

    QStringList loadReplyPool(const QString& appDir);
    bool waitForPageLoad(class WebDriverClient* client, const QString& expectedUrl = {});
    bool processDetailPage(class WebDriverClient* client);
    QString randomChoice(const QStringList& pool, const QString& exclude = {});
    int randomDelayMs(int minSec, int maxSec);
    bool askConfirm(const QString& title, const QString& message);
    bool askYesNo(const QString& title, const QString& message);
};
