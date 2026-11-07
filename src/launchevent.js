/*
 * Udalosti Outlooku. Klasicky Outlook pro Windows nacita JEN tento soubor (bez HTML),
 * proto ho webpack sestavuje vcetne MSAL do jednoho souboru.
 * Vsechny console.log jdou do runtime logu Outlooku (ladeni).
 */
import { getProfileForEvent } from "./profile";
import { getFromAddress, isAllriskFrom, applySignature, removeOwnSignature, notify } from "./signature";
import { SHOW_SIGNIN_HINT } from "./config";

Office.onReady(function () {});

async function handle(event, kind) {
  try {
    const from = await getFromAddress();
    console.log("[allrisk-sig] " + kind + " from=" + from);

    if (!isAllriskFrom(from)) {
      // Jina domena: jen uklidit nas podpis, pokud tam z drivejska je
      if (kind === "fromChanged") await removeOwnSignature();
      return;
    }

    let result;
    try {
      result = await getProfileForEvent();
    } catch (e) {
      console.log("[allrisk-sig] profil nelze nacist: " + e);
      // Podpis nevlozime -> neni hlavicka -> podpis doda transport rule (pokud existuje)
      if (SHOW_SIGNIN_HINT) await notify("Podpis Allrisk se nenačetl. Otevřete panel a přihlaste se.", true);
      return;
    }
    console.log("[allrisk-sig] profil z: " + result.source);

    if (!result.profile.firstName && !result.profile.lastName && !result.profile.displayName) {
      console.log("[allrisk-sig] profil bez jmena - podpis se nevklada");
      return;
    }
    await applySignature(result.profile, from);
  } catch (e) {
    console.log("[allrisk-sig] chyba: " + e);
  } finally {
    event.completed();
  }
}

function onNewMessageCompose(event) {
  handle(event, "compose");
}

function onMessageFromChanged(event) {
  handle(event, "fromChanged");
}

Office.actions.associate("onNewMessageCompose", onNewMessageCompose);
Office.actions.associate("onMessageFromChanged", onMessageFromChanged);
