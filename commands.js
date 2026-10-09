/* global Office */

// Menüdeki seçenekler: fonksiyon adı -> maile eklenecek metin
const METINLER = {
  biletiKapat: "Bileti kapat",
  biletiKapatma: "Bileti kapatma"
};

// Butonu kullanamayacak kişiler: fonksiyon adı -> engellenen e-posta adresleri (küçük harfle)
const ENGELLI_KULLANICILAR = {
  biletiKapat: ["okan.akdan@atptech.com"],
  biletiKapatma: []
};

function engelliMi(fonksiyonAdi) {
  const eposta = (Office.context.mailbox.userProfile.emailAddress || "").toLowerCase();
  return (ENGELLI_KULLANICILAR[fonksiyonAdi] || []).includes(eposta);
}

// Daha önce eklenmiş metni bulmak için ("Bileti kapat", "Bileti kapatma"nın içinde geçtiği için ayrı kontrol)
const MEVCUT_KONTROL = [
  { metin: "Bileti kapatma", desen: /Bileti kapatma/i },
  { metin: "Bileti kapat", desen: /Bileti kapat(?!ma)/i }
];

// Yanıtlarda alıntılanan eski yazışmanın başladığı yeri gösteren işaretler
const ALINTI_ISARETLERI = [
  /<div[^>]*id=["']?appendonsend["']?[^>]*>/i,                       // Outlook Web / yeni Outlook
  /<div[^>]*id=["']?divRplyFwdMsg["']?[^>]*>/i,                      // Outlook Web / yeni Outlook
  /<hr[^>]*display:\s*inline-block[^>]*>/i,                          // Outlook Web ayraç çizgisi
  /<div[^>]*border-top:\s*solid\s+#(E1E1E1|B5C4DF)[^>]*>/i,          // Klasik Windows Outlook
  /<div[^>]*class=["']?gmail_quote/i,                                // Gmail'den gelen yazışmalar
  /<blockquote/i                                                     // Genel alıntı bloğu
];

function alintiBaslangici(html) {
  let enKucuk = -1;
  for (const desen of ALINTI_ISARETLERI) {
    const m = desen.exec(html);
    if (m && (enKucuk === -1 || m.index < enKucuk)) enKucuk = m.index;
  }
  return enKucuk;
}

function bildir(item, mesaj, hata) {
  item.notificationMessages.replaceAsync("biletiKapat", {
    type: hata
      ? Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage
      : Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: mesaj,
    icon: "Icon.16",
    persistent: false
  });
}

function metinEkle(metin, event) {
  const item = Office.context.mailbox.item;
  const eklenecekHtml =
    '<p style="margin:0;font-family:Calibri,Arial,sans-serif;font-size:11pt;"><b>' +
    metin + "</b></p><br>";

  item.body.getAsync(Office.CoercionType.Html, (sonuc) => {
    if (sonuc.status !== Office.AsyncResultStatus.Succeeded) {
      bildir(item, "Mail gövdesi okunamadı: " + sonuc.error.message, true);
      event.completed();
      return;
    }

    const html = sonuc.value;
    const kesim = alintiBaslangici(html);
    let yazilanKisim, yeniHtml;

    if (kesim >= 0) {
      yazilanKisim = html.substring(0, kesim);
      yeniHtml = yazilanKisim + eklenecekHtml + html.substring(kesim);
    } else {
      const bodyKapanis = html.search(/<\/body>/i);
      yazilanKisim = bodyKapanis >= 0 ? html.substring(0, bodyKapanis) : html;
      yeniHtml = bodyKapanis >= 0
        ? yazilanKisim + eklenecekHtml + html.substring(bodyKapanis)
        : html + eklenecekHtml;
    }

    // Seçeneklerden biri zaten eklenmişse tekrar ekleme
    const duzMetin = yazilanKisim.replace(/<[^>]+>/g, "").replace(/&nbsp;/g, " ");
    for (const k of MEVCUT_KONTROL) {
      if (k.desen.test(duzMetin)) {
        bildir(item, "'" + k.metin + "' zaten eklenmiş. Değiştirmek için önce mevcut metni silin.");
        event.completed();
        return;
      }
    }

    item.body.setAsync(yeniHtml, { coercionType: Office.CoercionType.Html }, (yaz) => {
      if (yaz.status !== Office.AsyncResultStatus.Succeeded) {
        bildir(item, "Metin eklenemedi: " + yaz.error.message, true);
      }
      event.completed();
    });
  });
}

// Ekranın ortasında "Tamam" butonlu uyarı penceresi açar
function uyariPenceresi(mesaj, event) {
  const adres = new URL("uyari.html", window.location.href);
  adres.searchParams.set("mesaj", mesaj);

  Office.context.ui.displayDialogAsync(
    adres.toString(),
    { height: 30, width: 30, displayInIframe: true },
    (sonuc) => {
      if (sonuc.status !== Office.AsyncResultStatus.Succeeded) {
        // Pencere açılamazsa şerit uyarıya geri dön
        bildir(Office.context.mailbox.item, mesaj, true);
        event.completed();
        return;
      }
      const pencere = sonuc.value;
      const kapat = () => { try { pencere.close(); } catch (e) {} event.completed(); };
      pencere.addEventHandler(Office.EventType.DialogMessageReceived, kapat);
      pencere.addEventHandler(Office.EventType.DialogEventReceived, () => event.completed());
    }
  );
}

function calistir(fonksiyonAdi, event) {
  if (engelliMi(fonksiyonAdi)) {
    uyariPenceresi("Bu butonu kullanma yetkiniz yok.", event);
    return;
  }
  // GEÇİCİ TEŞHİS: Outlook'un gördüğü e-posta adresini gösterir (sorun çözülünce bu satırı silin)
  bildir(Office.context.mailbox.item, "Teşhis - algılanan adres: " +
    Office.context.mailbox.userProfile.emailAddress + " | sürüm: v6");
  metinEkle(METINLER[fonksiyonAdi], event);
}

function biletiKapat(event)   { calistir("biletiKapat", event); }
function biletiKapatma(event) { calistir("biletiKapatma", event); }

// XML manifest'teki <FunctionName> değerleriyle eşleşmesi için global tanım
window.biletiKapat = biletiKapat;
window.biletiKapatma = biletiKapatma;

Office.onReady(() => {
  if (Office.actions && Office.actions.associate) {
    Office.actions.associate("biletiKapat", biletiKapat);
    Office.actions.associate("biletiKapatma", biletiKapatma);
  }
});
