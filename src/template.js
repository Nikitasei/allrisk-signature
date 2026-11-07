/*
 * Sablona podpisu Allrisk - prevzata z AllriskTemplate.html (puvodni PowerShell skript).
 */

const WEB = "https://www.allrisk.cz";
const DEFAULT_COMPANY = "Allrisk, a.s.";
// Vychozi adresa, kdyz ji uzivatel nema vyplnenou v Entra ID
const DEFAULT_ADDRESS = { street: "Komárovská 263/20a", postalCode: "617 00", city: "Brno" };

// Fotka pro lidi, kteri nemaji URL v extensionAttribute15
const DEFAULT_PHOTO = "https://www.allrisk.cz/wp-content/uploads/2024/12/kulicka_fotka_email.png";

const LOGO = { src: "https://www.allrisk.cz/wp-content/uploads/2025/02/logo_allrisk_new_podpis_2025.png", width: 150, alt: "Allrisk" };

const SOCIAL = [
  { href: "https://www.facebook.com/allriskcz", icon: "https://www.allrisk.cz/wp-content/uploads/2022/08/facebook_podpis.png", alt: "Facebook" },
  { href: "https://www.linkedin.com/company/allriskgroup", icon: "https://www.allrisk.cz/wp-content/uploads/2022/08/linkedin_podpis.png", alt: "LinkedIn" },
  { href: "https://www.youtube.com/@allrisk.a.s.", icon: "https://www.allrisk.cz/wp-content/uploads/2022/08/youtube_podpis.png", alt: "YouTube" }
];

// Banner pod podpisem. null = bez banneru.
const BANNER = {
  src: "https://www.allrisk.cz/wp-content/uploads/2026/11/email-banner_420x200-1a-1.png",
  href: "https://energie.allrisk.cz/",
  width: 480,
  height: 200,
  alt: "Allrisk Energie"
};

/* ---------------------------------------------------------------- */

export function esc(v) {
  return String(v == null ? "" : v)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

const FONT = "font-family:tahoma,sans-serif;";

function div(style, content) {
  return '<div style="' + FONT + style + '">' + content + "</div>";
}

/** Jmeno: DisplayName jako v puvodnim skriptu; kdyz chybi, krestni + prijmeni. */
export function signatureName(p) {
  return p.displayName || [p.firstName, p.lastName].filter(Boolean).join(" ");
}

/**
 * @param {object} p  normalizovany profil (viz profile.js)
 * @param {string} email  e-mail, ktery se ukaze v podpisu
 */
export function buildSignature(p, email) {
  const name = esc(signatureName(p));
  const street = p.street || DEFAULT_ADDRESS.street;
  const postal = p.street ? p.postalCode : DEFAULT_ADDRESS.postalCode;
  const city = p.street ? p.city : DEFAULT_ADDRESS.city;
  const company = p.company || DEFAULT_COMPANY;

  let h = "<div>";
  h += "<strong>S pozdravem</strong>";
  h += '<table style="padding:0px;margin:10px 0;border:none;max-width:560px;"><tbody><tr>';

  // Levy horni: fotka (odkaz na web)
  h += '<td style="vertical-align:top;padding:0 16px 0 0;width:170px;">';
  h += '<a href="' + WEB + '" style="text-decoration:none;" target="_blank"><img src="' + esc(p.photo || DEFAULT_PHOTO) +
    '" width="170" alt="Fotka" style="border:0;"></a>';
  h += "</td>";

  // Pravy sloupec pres oba radky: udaje
  h += '<td rowspan="2" style="vertical-align:top;border-left:2px solid #e8e8e8;padding:0 0 0 16px;max-width:480px;">';
  h += div("font-size:14px;line-height:16px;font-weight:bold;color:black;", name);
  if (p.title) h += div("font-size:14px;line-height:16px;font-weight:normal;color:black;margin-bottom:10px;", esc(p.title));
  if (p.phone) {
    h += div("line-height:14px;font-weight:normal;color:black;margin-bottom:4px;",
      '<span style="font-size:14px;text-decoration:none;color:#134865;"><strong>' + esc(p.phone) + "</strong>&nbsp;</span>");
  }
  h += div("line-height:14px;font-weight:normal;color:black;margin-bottom:12px;",
    '<span style="font-size:14px;"><strong><a style="text-decoration:none;color:#134865;" href="mailto:' + esc(email) +
    '" target="_blank">' + esc(email) + "</a></strong>&nbsp;</span>");
  h += div("line-height:16px;font-weight:bold;color:black;", '<span style="font-size:14px;">' + esc(company) + "</span>");
  h += div("line-height:16px;font-weight:normal;color:black;margin-bottom:12px;",
    '<span style="font-size:14px;">' + esc(street) + "<br>" + esc(postal) + " " + esc(city) + "<br></span><br>" +
    '<a href="' + WEB + '" target="_blank">www.allrisk.cz</a><br>');
  h += SOCIAL.map(function (s) {
    return '<a href="' + esc(s.href) + '" style="text-decoration:none;" target="_blank"><img src="' + esc(s.icon) +
      '" width="24" height="24" alt="' + esc(s.alt) + '" style="border:0;"></a>';
  }).join(" ");
  h += "</td></tr>";

  // Levy dolni: logo
  h += '<tr><td style="vertical-align:bottom;padding:0 16px 0 0;width:170px;">';
  h += '<img src="' + esc(LOGO.src) + '" width="' + LOGO.width + '" alt="' + esc(LOGO.alt) + '" style="border:0;">';
  h += "</td></tr></tbody></table>";

  if (BANNER) {
    h += '<a href="' + esc(BANNER.href) + '" target="_blank"><img src="' + esc(BANNER.src) + '" width="' + BANNER.width +
      '" height="' + BANNER.height + '" alt="' + esc(BANNER.alt) + '" style="border:0;"></a>';
  }
  h += "</div>";
  return h;
}
