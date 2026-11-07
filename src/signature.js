/*
 * Rozhodnuti, jaky podpis vlozit, a prace s rozepsanou zpravou.
 */
import { DOMAINS, SIG_HEADER } from "./config";
import { buildSignature } from "./template";

export function domainOf(addr) {
  return (String(addr || "").split("@")[1] || "").toLowerCase();
}

/** Adresa v poli Od (u sdilene schranky / aliasu se lisi od prihlaseneho uzivatele). */
export function getFromAddress() {
  return new Promise(function (resolve) {
    const item = Office.context.mailbox.item;
    const fallback = function () {
      resolve((Office.context.mailbox.userProfile.emailAddress || "").toLowerCase());
    };
    if (item.from && item.from.getAsync) {
      item.from.getAsync(function (r) {
        if (r.status === Office.AsyncResultStatus.Succeeded && r.value && r.value.emailAddress) {
          resolve(r.value.emailAddress.toLowerCase());
        } else {
          fallback();
        }
      });
    } else {
      fallback();
    }
  });
}

/** Patri adresa Od do domen Allrisk? */
export function isAllriskFrom(from) {
  return DOMAINS.indexOf(domainOf(from)) !== -1;
}

/**
 * Jaky e-mail ukazat v podpisu:
 * - pise ze sve adresy / aliasu -> ta adresa
 * - pise ze sdilene schranky -> svuj osobni e-mail
 */
export function emailForSignature(profile, from) {
  if (profile.addresses && profile.addresses.indexOf(from) !== -1) return from;
  return profile.mail || from;
}

function call(fn) {
  return new Promise(function (resolve) { fn(resolve); });
}

export async function applySignature(profile, from) {
  const item = Office.context.mailbox.item;
  const html = buildSignature(profile, emailForSignature(profile, from));
  // disableClientSignatureAsync se zamerne NEPOUZIVA: v klasickem Outlooku a na Macu
  // by trvale vypnul vychozi podpis celeho uctu.
  const r = await call(function (cb) {
    item.body.setSignatureAsync(html, { coercionType: Office.CoercionType.Html }, cb);
  });
  if (r.status !== Office.AsyncResultStatus.Succeeded) {
    throw new Error("setSignatureAsync: " + (r.error && r.error.message));
  }
  if (item.internetHeaders && item.internetHeaders.setAsync) {
    const hdr = {};
    hdr[SIG_HEADER] = "addin";
    await call(function (cb) { item.internetHeaders.setAsync(hdr, cb); });
  }
}

/** Odebere podpis jen tehdy, kdyz ho vlozil tento doplnek. Cizi podpis nechava byt. */
export async function removeOwnSignature() {
  const item = Office.context.mailbox.item;
  if (!item.internetHeaders || !item.internetHeaders.getAsync) return;
  const r = await call(function (cb) { item.internetHeaders.getAsync([SIG_HEADER], cb); });
  const ours = r.status === Office.AsyncResultStatus.Succeeded && r.value && r.value[SIG_HEADER];
  if (!ours) return;
  await call(function (cb) { item.body.setSignatureAsync("", { coercionType: Office.CoercionType.Html }, cb); });
  await call(function (cb) { item.internetHeaders.removeAsync([SIG_HEADER], cb); });
}

export function notify(text, withSignInAction) {
  return new Promise(function (resolve) {
    try {
      const msg = {
        type: Office.MailboxEnums.ItemNotificationMessageType.InsightMessage,
        message: String(text).substring(0, 150),
        icon: "Icon16",
        actions: [{
          actionType: Office.MailboxEnums.ActionType.ShowTaskPane,
          actionText: "Otevřít",
          commandId: "allriskPaneBtn",
          contextData: "{}"
        }]
      };
      if (!withSignInAction) {
        msg.type = Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage;
        msg.persistent = false;
        delete msg.actions;
      }
      Office.context.mailbox.item.notificationMessages.replaceAsync("allriskSig", msg, function () { resolve(); });
    } catch (e) {
      resolve();
    }
  });
}
