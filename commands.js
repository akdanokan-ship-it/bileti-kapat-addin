/* global Office */

const METIN = "Bileti kapat";
const EKLENECEK_HTML =
  '<p style="margin:0;font-family:Calibri,Arial,sans-serif;font-size:11pt;"><b>' +
  METIN +
  "</b></p><br>";

// Yanıtlarda alıntılanan eski yazışmanın başladığı yeri gösteren işaretler.
// Metin, ilk bulunan işaretin hemen ÜSTÜNE, yani sizin yazdığınız kısmın sonuna eklenir.
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

function bildir(item, mesaj, tip) {
  item.notificationMessages.replaceAsync("biletiKapat", {
    type: tip || Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: mesaj,
    icon: "Icon.16",
    persistent: false
  });
}

function biletiKapat(event) {
  const item = Office.context.mailbox.item;

  item.body.getAsync(Office.CoercionType.Html, (sonuc) => {
    if (sonuc.status !== Office.AsyncResultStatus.Succeeded) {
      bildir(item, "Mail gövdesi okunamadı: " + sonuc.error.message,
        Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage);
      event.completed();
      return;
    }

    const html = sonuc.value;
    const kesim = alintiBaslangici(html);

    // Yanıt değilse (alıntı yoksa) gövdenin en sonuna, </body> öncesine ekle
    let yeniHtml;
    let yazilanKisim;
    if (kesim >= 0) {
      yazilanKisim = html.substring(0, kesim);
      yeniHtml = yazilanKisim + EKLENECEK_HTML + html.substring(kesim);
    } else {
      const bodyKapanis = html.search(/<\/body>/i);
      yazilanKisim = bodyKapanis >= 0 ? html.substring(0, bodyKapanis) : html;
      yeniHtml = bodyKapanis >= 0
        ? yazilanKisim + EKLENECEK_HTML + html.substring(bodyKapanis)
        : html + EKLENECEK_HTML;
    }

    // Butona iki kez basılırsa tekrar eklemesin
    if (yazilanKisim.replace(/<[^>]+>/g, "").includes(METIN)) {
      bildir(item, "'" + METIN + "' zaten eklenmiş.");
      event.completed();
      return;
    }

    item.body.setAsync(yeniHtml, { coercionType: Office.CoercionType.Html }, (yazSonuc) => {
      if (yazSonuc.status !== Office.AsyncResultStatus.Succeeded) {
        bildir(item, "Metin eklenemedi: " + yazSonuc.error.message,
          Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage);
      }
      event.completed();
    });
  });
}

// XML manifest'teki <FunctionName> ile eşleşmesi için global tanım
window.biletiKapat = biletiKapat;

Office.onReady(() => {
  if (Office.actions && Office.actions.associate) {
    Office.actions.associate("biletiKapat", biletiKapat);
  }
});
