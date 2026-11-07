/*
 * Prihlaseni (NAA) a nacteni profilu z Microsoft Graph.
 * Spolecne pro udalosti (launchevent.js) i panel (taskpane.js).
 */
import { createNestablePublicClientApplication, InteractionRequiredAuthError } from "@azure/msal-browser";
import { CLIENT_ID, TENANT_ID, PROFILE_MAX_AGE, PROFILE_FALLBACK_AGE } from "./config";

const SCOPES = ["User.Read"];
const SELECT = [
  "displayName", "givenName", "surname", "jobTitle", "businessPhones", "mobilePhone", "companyName",
  "mail", "userPrincipalName", "proxyAddresses", "streetAddress", "postalCode", "city",
  "onPremisesExtensionAttributes"
].join(",");
const CACHE_KEY = "allriskSignatureProfile";

let pcaPromise = null;

function getPca() {
  if (!pcaPromise) {
    pcaPromise = createNestablePublicClientApplication({
      auth: {
        clientId: CLIENT_ID,
        authority: "https://login.microsoftonline.com/" + TENANT_ID
      }
    });
  }
  return pcaPromise;
}

/** Token bez interakce. Vyhodi chybu, kdyz to nejde (udalosti nesmi otevirat okna). */
export async function getTokenSilent() {
  const pca = await getPca();
  const res = await pca.acquireTokenSilent({ scopes: SCOPES });
  return res.accessToken;
}

/** Token s moznosti prihlasovaciho okna - jen z panelu doplnku. */
export async function getTokenInteractive() {
  const pca = await getPca();
  try {
    return (await pca.acquireTokenSilent({ scopes: SCOPES })).accessToken;
  } catch (e) {
    if (e instanceof InteractionRequiredAuthError || (e && /interaction|consent|login/i.test(String(e.errorCode || e.message)))) {
      return (await pca.acquireTokenPopup({ scopes: SCOPES })).accessToken;
    }
    throw e;
  }
}

/** Prevede odpoved Graphu na to, co potrebuje sablona. */
export function normalize(g) {
  const ext = g.onPremisesExtensionAttributes || {};
  const addrs = [g.mail, g.userPrincipalName]
    .concat((g.proxyAddresses || []).filter(function (a) { return /^smtp:/i.test(a); }).map(function (a) { return a.slice(5); }))
    .filter(Boolean)
    .map(function (a) { return a.toLowerCase(); });
  return {
    displayName: g.displayName || "",
    firstName: g.givenName || "",
    lastName: g.surname || "",
    title: g.jobTitle || "",
    // stejne poradi jako puvodni skript: mobil, jinak pevna linka
    phone: g.mobilePhone || (g.businessPhones && g.businessPhones[0]) || "",
    company: g.companyName || "",
    mail: (g.mail || g.userPrincipalName || "").toLowerCase(),
    addresses: Array.from(new Set(addrs)),
    street: g.streetAddress || "",
    postalCode: g.postalCode || "",
    city: g.city || "",
    photo: ext.extensionAttribute15 || ""
  };
}

export async function fetchProfile(token) {
  const r = await fetch("https://graph.microsoft.com/v1.0/me?$select=" + SELECT, {
    headers: { Authorization: "Bearer " + token }
  });
  if (!r.ok) {
    throw new Error("Graph " + r.status + ": " + (await r.text()).substring(0, 200));
  }
  return normalize(await r.json());
}

/* ---- mezipamet v roamingSettings (uklada se do schranky uzivatele) ---- */

function readCache() {
  try {
    const c = Office.context.roamingSettings.get(CACHE_KEY);
    return c && c.t && c.p ? c : null;
  } catch (e) {
    return null;
  }
}

export function saveCache(profile) {
  return new Promise(function (resolve) {
    try {
      Office.context.roamingSettings.set(CACHE_KEY, { t: Date.now(), p: profile });
      Office.context.roamingSettings.saveAsync(function () { resolve(); });
    } catch (e) {
      resolve();
    }
  });
}

/**
 * Profil pro udalosti: cerstva mezipamet -> Graph -> starsi mezipamet jako zaloha.
 * Vraci { profile, source } nebo vyhodi chybu.
 */
export async function getProfileForEvent() {
  const cache = readCache();
  const age = cache ? Date.now() - cache.t : Infinity;
  if (cache && age < PROFILE_MAX_AGE) {
    return { profile: cache.p, source: "cache" };
  }
  try {
    const profile = await fetchProfile(await getTokenSilent());
    await saveCache(profile);
    return { profile: profile, source: "graph" };
  } catch (e) {
    if (cache && age < PROFILE_FALLBACK_AGE) {
      return { profile: cache.p, source: "stale-cache" };
    }
    throw e;
  }
}
