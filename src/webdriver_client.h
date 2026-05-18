#pragma once
#include <QObject>
#include <QNetworkAccessManager>
#include <QNetworkReply>
#include <QJsonObject>
#include <QJsonArray>
#include <QJsonDocument>
#include <QEventLoop>
#include <QTimer>
#include <QThread>
#include <QUrl>
#include <QVector>

const QString ELEMENT_KEY = QStringLiteral("element-6066-11e4-a52e-4f735466cecf");

struct ElementRef {
    QJsonObject ref;

    bool valid() const {
        return !ref.value(ELEMENT_KEY).toString().isEmpty();
    }

    QString guid() const {
        return ref.value(ELEMENT_KEY).toString();
    }

    static ElementRef fromResponse(const QJsonObject& findResponse) {
        ElementRef e;
        QJsonValue val = findResponse.value("value");
        if (val.isObject())
            e.ref = val.toObject();
        return e;
    }

    static QVector<ElementRef> fromArrayResponse(const QJsonObject& findResponse) {
        QVector<ElementRef> result;
        QJsonValue val = findResponse.value("value");
        if (val.isArray()) {
            for (const auto& item : val.toArray()) {
                ElementRef e;
                e.ref = item.toObject();
                if (e.valid())
                    result.append(e);
            }
        }
        return result;
    }

    static ElementRef invalid() { return {}; }
};

class WebDriverClient : public QObject {
    Q_OBJECT
public:
    explicit WebDriverClient(QObject* parent = nullptr);
    ~WebDriverClient();

    void setBaseUrl(const QString& url);

    bool createSession(const QJsonObject& capabilities);
    void closeSession();
    bool isOpen() const;

    bool navigateTo(const QString& url);
    bool refresh();
    QString currentUrl();

    ElementRef findElement(const QString& cssSelector);
    QVector<ElementRef> findElements(const QString& cssSelector);
    QVector<ElementRef> findChildElements(const ElementRef& parent, const QString& cssSelector);

    QString elementText(const ElementRef& elem);
    void elementClick(const ElementRef& elem);
    void elementSendKeys(const ElementRef& elem, const QString& text);
    bool elementDisplayed(const ElementRef& elem);
    QString elementAttribute(const ElementRef& elem, const QString& name);

    QJsonValue executeScript(const QString& script, const QJsonArray& args = {});

    QString getCurrentWindowHandle();
    void switchToWindow(const QString& handle);
    QString newTab();
    void closeCurrentWindow();

    bool waitForUrlContains(const QString& substring, int timeoutMs = 15000);
    bool waitForElement(const QString& cssSelector, int timeoutMs = 15000);

private:
    QNetworkAccessManager* m_manager = nullptr;
    QString m_baseUrl;
    QString m_sessionId;

    QString sessionPrefix() const;
    QJsonObject sendRequest(const QByteArray& method, const QString& endpoint,
                            const QByteArray& bodyData = {}, bool hasBody = false);
    QJsonObject doGet(const QString& endpoint);
    QJsonObject doPost(const QString& endpoint, const QJsonObject& body = {});
    QJsonObject doDelete(const QString& endpoint);
    bool hasError(const QJsonObject& resp, QString* msg = nullptr) const;
    static void msleep(int ms) { QThread::msleep(ms); }
};
