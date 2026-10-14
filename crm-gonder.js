/* global Office */
// Okuma ekranında açık olan mailin tüm bilgilerini JSON olarak flow'a gönderir.
// Flow adresi ayarlar.json içindeki "flowUrl" alanından okunur.

var AYAR_ADRESI = "ayarlar.json";

// Her Office çağrısı en fazla 5 sn bekler; cevap gelmezse null ile devam eder
function guvenli(fn, ad) {
  return new Promise(function (coz) {
    var bitti = false;
    var sure = setTimeout(function () { if (!bitti) { bitti = true; ZAMAN_ASIMI.push(ad || "?"); coz(null); } }, 5000);
    try {
      fn(function (r) {
        if (bitti) return; bitti = true; clearTimeout(sure);
        coz(r && r.status === Office.AsyncResultStatus.Succeeded ? r.value : null);
      });
    } catch (e) { if (!bitti) { bitti = true; clearTimeout(sure); coz(null); } }
  });
}
var ZAMAN_ASIMI = [];

function zamanAsimli(promise, ms, ad) {
  return Promise.race([promise, new Promise(function (_, red) { setTimeout(function () { red(new Error(ad + " " + (ms / 1000) + " sn içinde cevap vermedi")); }, ms); })]);
}

function kisi(r) { return r ? { ad: r.displayName, eposta: r.emailAddress } : null; }
function kisiler(l) { return (l || []).map(kisi); }
function tarih(d) { return d ? new Date(d).toISOString() : null; }

function bildir(mesaj, hata) {
  Office.context.mailbox.item.notificationMessages.replaceAsync("crmGonder", {
    type: hata ? Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage
               : Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: mesaj.substring(0, 150),
    icon: "Icon.16",
    persistent: false
  });
}

function mailBilgileriniTopla() {
  var item = Office.context.mailbox.item;
  return Promise.all([
    guvenli(function (cb) { item.body.getAsync(Office.CoercionType.Text, cb); }, "gövde-metin"),
    guvenli(function (cb) { item.body.getAsync(Office.CoercionType.Html, cb); }, "gövde-html"),
    item.getAllInternetHeadersAsync ? guvenli(function (cb) { item.getAllInternetHeadersAsync(cb); }, "başlıklar") : null,
    item.categories ? guvenli(function (cb) { item.categories.getAsync(cb); }, "kategoriler") : null
  ]).then(function (s) {
    var up = Office.context.mailbox.userProfile;
    var diag = Office.context.mailbox.diagnostics || {};
    return {
      gonderimZamani: new Date().toISOString(),
      kullanici: { ad: up.displayName, eposta: up.emailAddress, saatDilimi: up.timeZone },
      istemci: { uygulama: diag.hostName, surum: diag.hostVersion },

      itemId: item.itemId || null,
      internetMessageId: item.internetMessageId || null,
      konusmaId: item.conversationId || null,
      itemClass: item.itemClass || null,

      konu: item.subject || null,
      normalizeKonu: item.normalizedSubject || null,
      from: kisi(item.from),
      sender: kisi(item.sender),
      to: kisiler(item.to),
      cc: kisiler(item.cc),

      olusturmaTarihi: tarih(item.dateTimeCreated),
      degisiklikTarihi: tarih(item.dateTimeModified),

      kategoriler: (s[3] || []).map(function (c) { return c.displayName; }),
      ekler: (item.attachments || []).map(function (a) {
        return { id: a.id, ad: a.name, tur: a.contentType, boyut: a.size, satirIci: a.isInline, ekTuru: a.attachmentType };
      }),

      govdeMetin: s[0],
      govdeHtml: s[1],
      internetBasliklari: s[2]
    };
  });
}

function crmeGonder(event) {
  ZAMAN_ASIMI = [];
  bildir("1/4 Ayarlar okunuyor…");
  zamanAsimli(fetch(AYAR_ADRESI + "?t=" + Date.now(), { cache: "no-store" }), 10000, "ayarlar.json")
    .then(function (r) { return r.json(); })
    .then(function (ayar) {
      if (!ayar.flowUrl || ayar.flowUrl.indexOf("http") !== 0) throw new Error("ayarlar.json içinde flowUrl tanımlı değil");
      bildir("2/4 Mail bilgileri toplanıyor…");
      return mailBilgileriniTopla().then(function (veri) {
        var govde = JSON.stringify(veri);
        bildir("3/4 Flow'a gönderiliyor (" + Math.round(govde.length / 1024) + " KB)" +
               (ZAMAN_ASIMI.length ? " | yanıtsız: " + ZAMAN_ASIMI.join(",") : "") + "…");
        return zamanAsimli(fetch(ayar.flowUrl, { method: "POST", headers: { "Content-Type": "application/json" }, body: govde }), 30000, "Flow");
      });
    })
    .then(function (r) {
      if (!r.ok) throw new Error("Flow hata döndü: " + r.status);
      return r.json().catch(function () { return {}; });
    })
    .then(function (cevap) {
      bildir("4/4 CRM'e gönderildi." + (cevap && cevap.kayitId ? " Kayıt: " + cevap.kayitId : ""));
      event.completed();
    })
    .catch(function (e) {
      bildir("Gönderilemedi: " + e.message, true);
      event.completed();
    });
}

window.crmeGonder = crmeGonder;
Office.onReady(function () {
  if (Office.actions && Office.actions.associate) Office.actions.associate("crmeGonder", crmeGonder);
});
