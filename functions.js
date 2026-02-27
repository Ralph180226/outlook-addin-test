// Logger
function log(msg){ console.log(msg); const box=document.getElementById("status"); if(box) box.textContent+=msg+"\n"; }

// Helpers
function isCompose(){ const it=Office.context.mailbox?.item; return !!(it && typeof it.addItemAttachmentAsync==="function"); }

// READ: formulier → open compose
function sendPhishingReport(){
  const item=Office.context.mailbox?.item;
  if(!item || !item.itemId){ log("Geen itemId gevonden."); return; }

  const id=item.itemId;
  const extra=document.getElementById("extra")?.value || "";

  const s=Office.context.roamingSettings;
  s.set("origMailId", id);
  s.set("extraText", extra);

  s.saveAsync(()=> {
    log("Gegevens opgeslagen. Open compose...");
    Office.context.mailbox.displayNewMessageForm({
      toRecipients: ["ondersteuning@itssunday.nl"],
      subject: "Phishingmelding ITS Sunday",
      htmlBody:
        "<p>Er is een phishingmelding verstuurd.</p>" +
        "<p><b>Extra info:</b><br>" + (extra.trim()?extra:"(geen)") + "</p>"
    });
  });
}

// COMPOSE: bijlage toevoegen
function attachOriginalMail(done){
  if(!isCompose()){ done?.(); return; }
  const s=Office.context.roamingSettings;
  const id=s.get("origMailId");
  if(!id){ log("Geen originele mail-id."); done?.(); return; }

  log("Originele e‑mail wordt toegevoegd...");
  Office.context.mailbox.item.addItemAttachmentAsync(id, "Originele-e-mail.eml", (res)=>{
    if(res.status===Office.AsyncResultStatus.Succeeded){
      log("Originele e‑mail toegevoegd ✓");
      s.remove("origMailId"); s.remove("extraText"); s.saveAsync();
    } else {
      log("Bijlage fout: "+JSON.stringify(res.error));
    }
    done?.();
  });
}

// Compose readiness
function waitForComposeReady(cb){
  const it=Office.context.mailbox?.item;
  if(it && typeof it.addItemAttachmentAsync==="function"){ cb(); return; }
  setTimeout(()=>waitForComposeReady(cb), 300);
}

// Auto‑run in compose pane (optioneel)
Office.onReady(()=>{ if(isCompose()) waitForComposeReady(()=>attachOriginalMail()); });

// Ribbon command → koppelen (Compose‑knop)
function attachOriginalMailCommand(event){
  waitForComposeReady(()=>attachOriginalMail(()=>event.completed()));
}
