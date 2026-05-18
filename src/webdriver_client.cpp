#include "webdriver_client.h"
#include <QDateTime>

// ─── Static constants ────────────────────────────────────────────

namespace {
    const int REQUEST_TIMEOUT_MS = 30000;
    const int POLL_INTERVAL_MS = 500;
}

// ─── Constructor / Destructor ────────────────────────────────────

WebDriverClient::WebDriverClient(QObject* parent)
    : QObject(parent)
    , m_manager(new QNetworkAccessManager(this))
{
}

WebDriverClient::~WebDriverClient()
{
    if (isOpen())
        closeSession();
}

// ─── Session management ──────────────────────────────────────────

void WebDriverClient::setBaseUrl(const QString& url)
{
    m_baseUrl = url;
}

bool WebDriverClient::createSession(const QJsonObject& capabilities)
{
    QJsonObject body;
    body["capabilities"] = capabilities;

    auto resp = doPost("/session", body);

    QString err;
    if (hasError(resp, &err)) {
        m_sessionId.clear();
        return false;
    }

    QJsonObject value = resp.value("value").toObject();
    m_sessionId = value.value("sessionId").toString();
    return !m_sessionId.isEmpty();
}

void WebDriverClient::closeSession()
{
    if (!m_sessionId.isEmpty()) {
        doDelete(sessionPrefix());
        m_sessionId.clear();
    }
}

bool WebDriverClient::isOpen() const
{
    return !m_sessionId.isEmpty();
}

QString WebDriverClient::sessionPrefix() const
{
    return "/session/" + m_sessionId;
}

// ─── Navigation ──────────────────────────────────────────────────

bool WebDriverClient::navigateTo(const QString& url)
{
    QJsonObject body;
    body["url"] = url;
    auto resp = doPost(sessionPrefix() + "/url", body);
    QString err;
    return !hasError(resp, &err);
}

bool WebDriverClient::refresh()
{
    auto resp = doPost(sessionPrefix() + "/refresh");
    QString err;
    return !hasError(resp, &err);
}

QString WebDriverClient::currentUrl()
{
    auto resp = doGet(sessionPrefix() + "/url");
    return resp.value("value").toString();
}

// ─── Element finding ─────────────────────────────────────────────

ElementRef WebDriverClient::findElement(const QString& cssSelector)
{
    QJsonObject body;
    body["using"] = "css selector";
    body["value"] = cssSelector;

    auto resp = doPost(sessionPrefix() + "/element", body);
    return ElementRef::fromResponse(resp);
}

QVector<ElementRef> WebDriverClient::findElements(const QString& cssSelector)
{
    QJsonObject body;
    body["using"] = "css selector";
    body["value"] = cssSelector;

    auto resp = doPost(sessionPrefix() + "/elements", body);
    return ElementRef::fromArrayResponse(resp);
}

QVector<ElementRef> WebDriverClient::findChildElements(const ElementRef& parent, const QString& cssSelector)
{
    if (!parent.valid()) return {};

    QJsonObject body;
    body["using"] = "css selector";
    body["value"] = cssSelector;

    auto resp = doPost(sessionPrefix() + "/element/" + parent.guid() + "/elements", body);
    return ElementRef::fromArrayResponse(resp);
}

// ─── Element operations ──────────────────────────────────────────

QString WebDriverClient::elementText(const ElementRef& elem)
{
    if (!elem.valid()) return {};
    auto resp = doGet(sessionPrefix() + "/element/" + elem.guid() + "/text");
    return resp.value("value").toString();
}

void WebDriverClient::elementClick(const ElementRef& elem)
{
    if (!elem.valid()) return;
    doPost(sessionPrefix() + "/element/" + elem.guid() + "/click");
}

void WebDriverClient::elementSendKeys(const ElementRef& elem, const QString& text)
{
    if (!elem.valid()) return;
    QJsonObject body;
    body["text"] = text;
    doPost(sessionPrefix() + "/element/" + elem.guid() + "/value", body);
}

bool WebDriverClient::elementDisplayed(const ElementRef& elem)
{
    if (!elem.valid()) return false;
    auto resp = doGet(sessionPrefix() + "/element/" + elem.guid() + "/displayed");
    QString err;
    if (hasError(resp, &err)) return false;
    return resp.value("value").toBool(false);
}

QString WebDriverClient::elementAttribute(const ElementRef& elem, const QString& name)
{
    if (!elem.valid()) return {};
    auto resp = doGet(sessionPrefix() + "/element/" + elem.guid() + "/attribute/" + name);
    return resp.value("value").toString();
}

// ─── JavaScript ──────────────────────────────────────────────────

QJsonValue WebDriverClient::executeScript(const QString& script, const QJsonArray& args)
{
    QJsonObject body;
    body["script"] = script;
    body["args"] = args;
    auto resp = doPost(sessionPrefix() + "/execute/sync", body);
    QString err;
    if (hasError(resp, &err))
        return {};
    return resp.value("value");
}

// ─── Window / Tab management ─────────────────────────────────────

QString WebDriverClient::getCurrentWindowHandle()
{
    auto resp = doGet(sessionPrefix() + "/window");
    return resp.value("value").toString();
}

void WebDriverClient::switchToWindow(const QString& handle)
{
    QJsonObject body;
    body["handle"] = handle;
    doPost(sessionPrefix() + "/window", body);
}

QString WebDriverClient::newTab()
{
    QJsonObject body;
    body["type"] = "tab";
    auto resp = doPost(sessionPrefix() + "/window/new", body);
    QJsonObject val = resp.value("value").toObject();
    return val.value("handle").toString();
}

void WebDriverClient::closeCurrentWindow()
{
    doDelete(sessionPrefix() + "/window");
}

// ─── Wait helpers ────────────────────────────────────────────────

bool WebDriverClient::waitForUrlContains(const QString& substring, int timeoutMs)
{
    qint64 deadline = QDateTime::currentMSecsSinceEpoch() + timeoutMs;
    while (QDateTime::currentMSecsSinceEpoch() < deadline) {
        if (currentUrl().contains(substring))
            return true;
        msleep(POLL_INTERVAL_MS);
    }
    return false;
}

bool WebDriverClient::waitForElement(const QString& cssSelector, int timeoutMs)
{
    qint64 deadline = QDateTime::currentMSecsSinceEpoch() + timeoutMs;
    while (QDateTime::currentMSecsSinceEpoch() < deadline) {
        auto elems = findElements(cssSelector);
        if (!elems.isEmpty())
            return true;
        msleep(POLL_INTERVAL_MS);
    }
    return false;
}

// ─── HTTP helpers (synchronous via nested event loop) ────────────

QJsonObject WebDriverClient::sendRequest(const QByteArray& method,
                                          const QString& endpoint,
                                          const QByteArray& bodyData,
                                          bool hasBody)
{
    QNetworkRequest request(QUrl(m_baseUrl + endpoint));
    request.setRawHeader("Accept", "application/json");

    QNetworkReply* reply = nullptr;
    if (method == "GET") {
        reply = m_manager->get(request);
    } else if (method == "DELETE") {
        reply = m_manager->deleteResource(request);
    } else {
        request.setHeader(QNetworkRequest::ContentTypeHeader, "application/json");
        reply = m_manager->post(request, bodyData);
    }

    QEventLoop loop;
    QTimer timer;
    timer.setSingleShot(true);
    QObject::connect(reply, &QNetworkReply::finished, &loop, &QEventLoop::quit);
    QObject::connect(&timer, &QTimer::timeout, &loop, &QEventLoop::quit);
    timer.start(REQUEST_TIMEOUT_MS);

    loop.exec();  // blocks until finished or timeout

    QJsonObject result;
    if (!reply->isFinished()) {
        reply->abort();
        result["value"] = QJsonObject{
            {"error", "timeout"},
            {"message", "WebDriver request timed out"}
        };
    } else if (reply->error() != QNetworkReply::NoError) {
        result["value"] = QJsonObject{
            {"error", reply->errorString()}
        };
    } else {
        QByteArray data = reply->readAll();
        if (!data.isEmpty()) {
            QJsonParseError parseErr;
            QJsonDocument doc = QJsonDocument::fromJson(data, &parseErr);
            if (parseErr.error == QJsonParseError::NoError)
                result = doc.object();
            else
                result["value"] = QJsonObject{{"error", parseErr.errorString()}};
        }
    }

    reply->deleteLater();
    return result;
}

QJsonObject WebDriverClient::doGet(const QString& endpoint)
{
    return sendRequest("GET", endpoint);
}

QJsonObject WebDriverClient::doPost(const QString& endpoint, const QJsonObject& body)
{
    return sendRequest("POST", endpoint, QJsonDocument(body).toJson(QJsonDocument::Compact), true);
}

QJsonObject WebDriverClient::doDelete(const QString& endpoint)
{
    return sendRequest("DELETE", endpoint);
}

bool WebDriverClient::hasError(const QJsonObject& resp, QString* msg) const
{
    QJsonValue val = resp.value("value");
    if (val.isObject()) {
        QJsonObject o = val.toObject();
        if (o.contains("error") && !o.value("error").toString().isEmpty()) {
            if (msg)
                *msg = o.value("error").toString() + ": " + o.value("message").toString();
            return true;
        }
    }
    return false;
}
