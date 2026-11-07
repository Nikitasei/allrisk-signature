/*
 * Panel doplnku: rucni prihlaseni (kdyz tiche selze), vlozeni podpisu a diagnostika pro pilot.
 */
import { getTokenInteractive, fetchProfile, saveCache } from "./profile";
import { getFromAddress, isAllriskFrom, applySignature, emailForSignature } from "./signature";

function $(id) { return document.getElementById(id); }

function status(text, cls) {
  const s = $("status");
  s.textContent = text;
  s.className = cls || "";
}

function show(rows) {
  const dl = $("details");
  dl.innerHTML = "";
  rows.forEach(function (r) {
    const dt = document.createElement("dt");
    dt.textContent = r[0];
    const dd = document.createElement("dd");
    dd.textContent = r[1] || "chybí";
    if (!r[1]) dd.className = "missing";
    dl.appendChild(dt);
    dl.appendChild(dd);
  });
}

async function run() {
  const btn = $("insert");
  btn.disabled = true;
  try {
    status("Přihlašuji…");
    const profile = await fetchProfile(await getTokenInteractive());
    await saveCache(profile);
    const from = await getFromAddress();
    show([
      ["Od", from],
      ["Jméno", [profile.firstName, profile.lastName].filter(Boolean).join(" ")],
      ["Funkce", profile.title],
      ["Telefon", profile.phone],
      ["E-mail v podpisu", emailForSignature(profile, from)],
      ["Adresa", [profile.street, profile.postalCode, profile.city].filter(Boolean).join(", ")],
      ["Fotka (ext.Attr15)", profile.photo]
    ]);
    if (!isAllriskFrom(from)) {
      status("Adresa Od není Allrisk – podpis se nevkládá.", "err");
      return;
    }
    await applySignature(profile, from);
    status("Podpis vložen. Příště už se vloží sám.", "ok");
  } catch (e) {
    status("Chyba: " + (e && e.message ? e.message : e), "err");
  } finally {
    btn.disabled = false;
  }
}

Office.onReady(function () {
  const btn = $("insert");
  btn.disabled = false;
  btn.addEventListener("click", run);
});
