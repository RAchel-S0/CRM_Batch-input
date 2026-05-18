#include "logic_handler.h"
#include "webdriver_client.h"

#include <QCoreApplication>
#include <QDir>
#include <QFile>
#include <QFileInfo>
#include <QProcess>
#include <QTextStream>
#include <QRandomGenerator>
#include <QThread>
#include <QDateTime>
#include <QJsonObject>
#include <QJsonArray>
#include <QJsonValue>
#include <QNetworkAccessManager>
#include <QNetworkReply>
#include <QNetworkRequest>
#include <QEventLoop>
#include <QTimer>
#include <QUrl>

// ─── Default reply pool ──────────────────────────────────────────

static const QStringList DEFAULT_POOL = {
    "推荐用户使用诺企服开票，用户已知晓",
    "客户反馈系统稳定，暂无新增需求",
    "电话无人接听，稍后再次联系",
    "已告知客户最新的优惠活动",
    "常规回访，客户表示满意"
};

// ─── CSS Selectors ───────────────────────────────────────────────

struct S {
    static constexpr const char* listRow     = "div[data-fieldname='name']";
    static constexpr const char* lastRecord  = ".fxeditor-render-text";
    static constexpr const char* publishBtn  = ".d-salelog_publish_btn";
    static constexpr const char* editor      = ".tiptap.ProseMirror";
    static constexpr const char* selectInput = "input.j-select-input";
    static constexpr const char* submitBtn   = "span.j-ok[action-type='dialogEnter']";
};

namespace WdKeys {
    const QString ARROW_DOWN("\uE015");
    const QString ENTER("\uE007");
}

// ─── Task struct ─────────────────────────────────────────────────

struct Task {
    QString name;
    QString url;
};

// ─── Constructor ─────────────────────────────────────────────────

LogicHandler::LogicHandler(QObject* parent)
    : QObject(parent)
{
    m_running.storeRelaxed(0);
}

// ─── Public slots ────────────────────────────────────────────────

void LogicHandler::doStop()
{
    m_running.storeRelaxed(0);
    emit log("正在停止... (请等待当前操作完成)");
}

void LogicHandler::doRun(const RunParams& params)
{
    m_running.storeRelaxed(1);

    // ── 1. Load reply pool ────────────────────────────────────
    emit log("──────────────────────────────");
    m_replyPool = loadReplyPool(params.appDir);
    emit log("──────────────────────────────");

    // ── 2. Start ChromeDriver ─────────────────────────────────
    QString driverExe = params.driverPath;
    if (!QFileInfo::exists(driverExe)) {
        emit log(QString("错误: 找不到 chromedriver.exe"));
        emit log(QString("期望路径: %1").arg(QDir::toNativeSeparators(driverExe)));
        emit log("请将 chromedriver.exe 与本程序放在同一目录");
        emit finished(0);
        return;
    }

    emit log("正在启动 ChromeDriver...");
    QProcess chromedriver;
    chromedriver.setProcessChannelMode(QProcess::MergedChannels);
    chromedriver.start(driverExe, {"--port=9515"});

    if (!chromedriver.waitForStarted(5000)) {
        emit log("错误: 无法启动 ChromeDriver");
        emit log("提示：可能端口 9515 被占用，请检查是否已有 chromedriver 在运行");
        emit finished(0);
        return;
    }

    // Wait for ChromeDriver HTTP status endpoint to respond
    {
        QNetworkAccessManager checkNet;
        bool ready = false;
        qint64 deadline = QDateTime::currentMSecsSinceEpoch() + 15000;
        while (QDateTime::currentMSecsSinceEpoch() < deadline) {
            if (chromedriver.state() != QProcess::Running) {
                emit log("错误: ChromeDriver 意外退出");
                QByteArray out = chromedriver.readAllStandardOutput();
                if (!out.isEmpty())
                    emit log(QString("输出: %1").arg(QString::fromUtf8(out).trimmed()));
                emit finished(0);
                return;
            }
            QNetworkRequest req(QUrl("http://localhost:9515/status"));
            QNetworkReply* reply = checkNet.get(req);
            QEventLoop loop;
            QTimer t;
            t.setSingleShot(true);
            QObject::connect(reply, &QNetworkReply::finished, &loop, &QEventLoop::quit);
            QObject::connect(&t, &QTimer::timeout, &loop, &QEventLoop::quit);
            t.start(2000);
            loop.exec();
            if (reply->error() == QNetworkReply::NoError) {
                ready = true;
                reply->deleteLater();
                break;
            }
            reply->deleteLater();
            QThread::msleep(500);
        }
        if (!ready) {
            emit log("错误: ChromeDriver 启动超时");
            emit log("提示: 请确认 chromedriver.exe 版本与 Chrome 浏览器版本匹配");
            chromedriver.kill();
            chromedriver.waitForFinished();
            emit finished(0);
            return;
        }
    }
    emit log("ChromeDriver 已就绪");

    // ── 3. Create WebDriver session ───────────────────────────
    WebDriverClient client;
    client.setBaseUrl("http://localhost:9515");

    // Ensure user data dir exists
    QString userDataDir = params.appDir + "/ChromeUserData";
    QDir().mkpath(userDataDir);

    QJsonObject chromeOpts;
    QJsonArray chromeArgs;
    chromeArgs.append("--start-maximized");
    chromeArgs.append("--ignore-certificate-errors");
    chromeArgs.append("--user-data-dir=" + QDir::toNativeSeparators(userDataDir));
    chromeOpts["args"] = chromeArgs;
    chromeOpts["excludeSwitches"] = QJsonArray{"enable-logging"};

    if (!params.chromePath.isEmpty() && QFileInfo::exists(params.chromePath)) {
        chromeOpts["binary"] = QDir::toNativeSeparators(params.chromePath);
    } else {
        emit log("错误：Chrome 路径不存在");
        chromedriver.kill();
        chromedriver.waitForFinished();
        emit finished(0);
        return;
    }

    QJsonObject caps;
    caps["browserName"] = "chrome";
    caps["goog:chromeOptions"] = chromeOpts;

    QJsonObject sessionCaps;
    sessionCaps["alwaysMatch"] = caps;

    if (!client.createSession(sessionCaps)) {
        emit log("启动浏览器失败，请先关闭所有已打开的 Chrome 窗口");
        chromedriver.kill();
        chromedriver.waitForFinished();
        emit finished(0);
        return;
    }
    emit log("浏览器已启动");

    // ── 4. Navigate to CRM ────────────────────────────────────
    emit log("正在打开目标网址...");
    client.navigateTo(params.url);

    // ── 5. User confirmation ──────────────────────────────────
    if (!askConfirm("准备就绪",
            "请确认：\n1. 已登录 CRM\n2. 已进入[客户-未回访]列表页\n\n点击【确定】开始任务。")) {
        client.closeSession();
        chromedriver.kill();
        chromedriver.waitForFinished();
        emit log("用户取消操作");
        emit finished(0);
        return;
    }

    // ── 6. Main processing loop ───────────────────────────────
    int successCount = 0;
    int pageRound = 0;

    while (successCount < params.limit && m_running.loadRelaxed()) {

        QString listHandle = client.getCurrentWindowHandle();
        emit log("正在扫描当前页...");
        QThread::msleep(2000);

        // ── Collect tasks from all visible rows via JS ───────
        // Matches Python: iterate ALL <a> for name, pick LASK <a[href*=AccountObj] for url
        QString collectJs =
            "var rows = document.querySelectorAll(arguments[0]);"
            "var tasks = [];"
            "for (var i = 0; i < rows.length; i++) {"
            "  var row = rows[i];"
            "  if (row.offsetParent === null) continue;"
            "  var name = '';"
            "  var url = '';"
            "  var allLinks = row.querySelectorAll('a');"
            "  for (var j = 0; j < allLinks.length; j++) {"
            "    var text = allLinks[j].textContent.trim();"
            "    if (text) name = text;"
            "    var href = allLinks[j].getAttribute('href');"
            "    if (href && href.indexOf('AccountObj') !== -1) {"
            "      url = href;"
            "    }"
            "  }"
            "  if (url) {"
            "    if (url.indexOf('://') === -1) url = 'https://www.fxiaoke.com' + url;"
            "    tasks.push({name: name || '未知客户', url: url});"
            "  }"
            "}"
            "return tasks;";

        QJsonArray args;
        args.append(QString(S::listRow));
        QJsonValue result = client.executeScript(collectJs, args);

        QVector<Task> pageTasks;
        for (const auto& v : result.toArray()) {
            QJsonObject o = v.toObject();
            QString url = o.value("url").toString();
            QString name = o.value("name").toString();
            if (!url.isEmpty()) {
                pageTasks.append({name, url});
            }
        }

        emit log(QString("当前页提取到 %1 个任务。").arg(pageTasks.size()));

        if (pageTasks.isEmpty()) {
            emit log(">>> 请手动翻页，等待加载完成后点击弹窗确认...");
            if (!askYesNo("翻页提示", "当前页无数据。\n请手动翻页后，点击【是】继续，【否】退出。"))
                break;
            continue;
        }

        // ── Process each task in a new tab ────────────────────
        emit log(">>> 开始处理 (您可以最小化浏览器)...");
        QString detailTabHandle = client.newTab();
        client.switchToWindow(detailTabHandle);

        for (int i = 0; i < pageTasks.size(); ++i) {
            if (successCount >= params.limit || !m_running.loadRelaxed())
                break;

            const auto& task = pageTasks[i];
            emit log(QString("[%1/%2] 处理: %3")
                         .arg(successCount + 1).arg(params.limit).arg(task.name));

            client.navigateTo(task.url);

            // Refresh to clear stale SPA DOM
            client.refresh();

            if (!waitForPageLoad(&client, task.url)) {
                emit log("   -> [跳过] 页面加载超时");
                continue;
            }

            if (processDetailPage(&client)) {
                emit log("   -> [成功]");
                ++successCount;
            } else {
                emit log("   -> [失败] 录入流程未完成");
            }

            int delayMs = randomDelayMs(params.minWait, params.maxWait);
            emit log(QString("   -> [休息] %1 秒...").arg(delayMs / 1000.0, 0, 'f', 1));
            QThread::msleep(delayMs);
        }

        // Close tab and switch back to list
        client.closeCurrentWindow();
        client.switchToWindow(listHandle);

        if (successCount >= params.limit || !m_running.loadRelaxed())
            break;

        // ── Pagination prompt ─────────────────────────────────
        if (!askYesNo("翻页提示",
                      QString("进度: %1/%3\n请手动翻页，完成后点击【是】继续。").arg(successCount).arg(params.limit)))
            break;
    }

    // ── 7. Clean up ───────────────────────────────────────────
    emit log("正在关闭浏览器...");
    client.closeSession();

    chromedriver.kill();
    chromedriver.waitForFinished(3000);

    emit log(QString("任务完成，共录入 %1 条。").arg(successCount));
    emit finished(successCount);
}

// ─── Reply pool loader ───────────────────────────────────────────

QStringList LogicHandler::loadReplyPool(const QString& appDir)
{
    QString configPath = appDir + "/reply_list.txt";
    QFile file(configPath);

    if (!file.exists()) {
        emit log(QString("未找到配置文件，生成默认文件: reply_list.txt"));
        if (file.open(QIODevice::WriteOnly | QIODevice::Text)) {
            QTextStream out(&file);
            for (const auto& line : DEFAULT_POOL)
                out << line << "\n";
            file.close();
        }
        return DEFAULT_POOL;
    }

    if (!file.open(QIODevice::ReadOnly | QIODevice::Text)) {
        emit log("配置文件读取失败，使用默认值");
        return DEFAULT_POOL;
    }

    QStringList pool;
    QTextStream in(&file);
    while (!in.atEnd()) {
        QString line = in.readLine().trimmed();
        if (!line.isEmpty())
            pool.append(line);
    }
    file.close();

    if (!pool.isEmpty()) {
        emit log(QString("成功加载话术库: %1 条").arg(pool.size()));
        return pool;
    }

    emit log("配置文件为空，使用默认值");
    return DEFAULT_POOL;
}

// ─── Page load wait ──────────────────────────────────────────────

bool LogicHandler::waitForPageLoad(WebDriverClient* client, const QString& expectedUrl)
{
    if (!expectedUrl.isEmpty()) {
        QString s = expectedUrl;
        int h = s.lastIndexOf('#');
        if (h >= 0) s = s.mid(h + 1);
        if (!client->waitForUrlContains(s, 15000))
            return false;
    } else {
        if (!client->waitForUrlContains("AccountObj", 15000))
            return false;
    }

    QThread::msleep(3000);
    return client->waitForElement(S::publishBtn, 15000);
}

// ─── Detail page processing ──────────────────────────────────────

bool LogicHandler::processDetailPage(WebDriverClient* client)
{
    // 1. Read last record ──────────────────────────────────────
    QString lastContent;
    QThread::msleep(2000);
    auto elements = client->findElements(S::lastRecord);
    for (const auto& el : elements) {
        QString t = client->elementText(el).trimmed();
        if (!t.isEmpty()) {
            lastContent = t;
            break;
        }
    }

    if (!lastContent.isEmpty())
        emit log(QString("   -> 历史: %1...").arg(lastContent.left(10)));
    else
        emit log("   -> 无历史记录");

    // 2. Pick content ──────────────────────────────────────────
    QString finalContent;
    QString cleanLast = lastContent.trimmed();
    bool isStandard = false;

    for (const auto& p : m_replyPool) {
        if (p.contains(cleanLast) || cleanLast.contains(p)) {
            isStandard = true;
            break;
        }
    }

    if (cleanLast.isEmpty() || isStandard) {
        finalContent = randomChoice(m_replyPool, cleanLast);
        emit log("   -> 策略: 随机库内容");
    } else {
        finalContent = cleanLast;
        emit log("   -> 策略: 复制上一条");
    }

    // 3. Click publish button ──────────────────────────────────
    if (!client->waitForElement(S::publishBtn, 15000)) {
        emit log("   -> 错误: 无法找到写跟进按钮");
        return false;
    }
    auto pubBtn = client->findElement(S::publishBtn);
    if (!pubBtn.valid()) {
        emit log("   -> 错误: 无法定位写跟进按钮");
        return false;
    }
    client->executeScript("arguments[0].click();", QJsonArray{pubBtn.ref});
    QThread::msleep(2000);

    // 4. Fill editor ───────────────────────────────────────────
    if (!client->waitForElement(S::editor, 15000)) {
        emit log("   -> 错误: 找不到输入框");
        return false;
    }
    auto editor = client->findElement(S::editor);
    if (!editor.valid()) {
        emit log("   -> 错误: 无法定位输入框");
        return false;
    }
    client->elementClick(editor);
    client->elementSendKeys(editor, finalContent);
    QThread::msleep(1000);

    // 5. Dropdown selection ────────────────────────────────────
    auto selects = client->findElements(S::selectInput);
    ElementRef target = ElementRef::invalid();
    if (selects.size() >= 2)
        target = selects[1];
    else if (selects.size() == 1)
        target = selects[0];

    if (target.valid()) {
        client->executeScript("arguments[0].click();", QJsonArray{target.ref});
        QThread::msleep(1000);
        client->elementSendKeys(target, "客户反馈");
        QThread::msleep(1500);
        client->elementSendKeys(target, WdKeys::ARROW_DOWN);
        QThread::msleep(500);
        client->elementSendKeys(target, WdKeys::ENTER);
        QThread::msleep(1000);
    } else {
        emit log("   -> 警告: 未找到下拉框");
    }

    // 6. Submit ────────────────────────────────────────────────
    auto submit = client->findElement(S::submitBtn);
    if (!submit.valid()) {
        emit log("   -> 错误: 提交按钮未找到");
        return false;
    }
    client->executeScript("arguments[0].click();", QJsonArray{submit.ref});

    for (int i = 0; i < 8; ++i) {
        QThread::msleep(1000);
        if (!client->elementDisplayed(submit))
            return true;
    }

    emit log("   -> 失败: 弹窗未关闭");
    return false;
}

// ─── Utilities ───────────────────────────────────────────────────

QString LogicHandler::randomChoice(const QStringList& pool, const QString& exclude)
{
    QStringList candidates;
    for (const auto& t : pool) {
        if (t != exclude)
            candidates.append(t);
    }
    if (candidates.isEmpty())
        candidates = pool;
    int idx = QRandomGenerator::global()->bounded(candidates.size());
    return candidates.at(idx);
}

int LogicHandler::randomDelayMs(int minSec, int maxSec)
{
    return QRandomGenerator::global()->bounded(minSec * 1000, maxSec * 1000 + 1);
}

bool LogicHandler::askConfirm(const QString& title, const QString& msg)
{
    bool result = false;
    emit needConfirm(title, msg, &result);
    return result;
}

bool LogicHandler::askYesNo(const QString& title, const QString& msg)
{
    bool result = false;
    emit needYesNo(title, msg, &result);
    return result;
}
