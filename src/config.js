/*
 * Jedine misto, ktere se pri nasazeni upravuje.
 */

// Z registrace aplikace v Entra ID (tenant allrisk.cz)
export const CLIENT_ID = "415a2379-08dc-44b0-bc18-a213fa2364e0";
export const TENANT_ID = "3ae8279c-8e16-4892-8e96-0f46badc9e48";

// Domeny v poli Od, pro ktere se vklada podpis Allrisk. Ostatni adresy doplnek nechava byt.
export const DOMAINS = ["allrisk.cz"];

// Hlavicka, podle ktere transport rule pozna, ze podpis uz vlozil doplnek
export const SIG_HEADER = "X-Allrisk-Signature";

// Jak dlouho se smi pouzit ulozeny profil, nez se znovu nacte z Graphu (ms)
export const PROFILE_MAX_AGE = 4 * 60 * 60 * 1000; // 4 hodiny
// Jak stary profil se smi pouzit jako zaloha, kdyz Graph nebo prihlaseni selze (ms)
export const PROFILE_FALLBACK_AGE = 30 * 24 * 60 * 60 * 1000; // 30 dni

// Pilot: kdyz se nepodari tiche prihlaseni, ukaze nad zpravou listu s odkazem na panel doplnku
export const SHOW_SIGNIN_HINT = true;
