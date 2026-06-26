# Aktualizace balíčků – návod k použití

Aplikace slouží produktovým manažerům (PM) k aktualizaci starých položek v produktových balíčcích a k přípravě importních souborů pro ERP.

---

## Co potřebujete ve složce

Tyto soubory musí ležet **ve stejné složce**:

- **`aktualizace_balicku.html`** – samotná aplikace (otevírá se dvojklikem v prohlížeči)
- **`katalog.js`** – databáze položek, ze které se vybírají náhrady (načítá se automaticky)
- **`balicky_sk.xlsx`** – soubor s balíčky, který do aplikace načtete

> Když budete aplikaci přesouvat nebo posílat, vždy berte `aktualizace_balicku.html` i `katalog.js` společně.

Aplikace nikdy nepřepisuje vaše vstupní soubory. Všechny výstupy ukládá jako **nové soubory**.

---

## 1. Spuštění a načtení balíčků

1. Zavřete `balicky_sk.xlsx` v Excelu (jinak ho prohlížeč nemůže otevřít – byl by zamčený).
2. Otevřete **`aktualizace_balicku.html`** dvojklikem.
3. Přetáhněte do okna `balicky_sk.xlsx` (nebo klikněte a vyberte ho).
4. Vyberte v seznamu **svého PM** a dejte *Pokračovat*. Zobrazí se jen vaše balíčky.

Katalog položek se načte sám na pozadí – nic nenačítáte ručně.

---

## 2. Seznam balíčků

Nahoře vidíte přehled: počet balíčků, **položek (můj PM)**, **položek (jiný PM)**, **starých položek k řešení**, upravených a smazaných balíčků.

U každého balíčku jsou štítky:

- **sdílený** – balíček obsahuje i položky jiného PM
- **upraveno** – v balíčku jste už něco změnil
- **N starých** / **hotovo** – kolik starých položek ještě čeká

Kliknutím na název balíček otevřete. Ikona **🗑** vpravo smaže celý balíček (viz část 6).

---

## 3. Práce v balíčku – aktualizace staré položky

Staré položky (názvy se ZZZ nebo !!!) jsou žlutě zvýrazněné štítkem **stará**. Pod každou se rovnou nabízejí **Navrhované** alternativy (modré štítky) podle skupiny výrobků, výrobce, PM a podobnosti názvu.

Máte tři možnosti:

- **Kliknout na navržený štítek** – nejrychlejší způsob, položku rovnou zamění.
- **Zaměnit** – otevře okno, kde nahoře vidíte návrhy a níže můžete **hledat fulltextem** (podle názvu nebo ID, bez ohledu na diakritiku). Vyberete položku a dáte *Potvrdit*.
- **Přidat novou** – přidá novou položku do balíčku, **stará zůstane** (použijte, když nejde o náhradu, ale o doplnění).
- **Smazat** – odebere položku z balíčku bez náhrady (maže rovnou, lze vrátit).

### Propagace záměny do dalších balíčků
V okně *Zaměnit* je zatržítko **„Při záměně nahradit tuto starou položku i ve všech ostatních balíčcích"** (výchozí zapnuto). Když je stejná stará položka i jinde, nahradí se všude najednou. Při rychlé záměně přes štítek se aplikace zeptá, jestli nahradit i v ostatních balíčcích, nebo jen v tomto.

### Přidání úplně nové položky
Nahoře v detailu balíčku je tlačítko **➕ Přidat položku**. Otevře vyhledávání v katalogu (fulltext podle názvu nebo ID), vyberete položku a *Potvrdíte* – přidá se do balíčku jako nová, nezávisle na starých položkách.

### Vrácení změny
Smazanou položku vrátíte tlačítkem **Vrátit** u daného řádku.

---

## 4. Sdílené balíčky (více PM)

Některé balíčky obsahují položky více PM. Protože ERP nahrazuje **celý balíček**, aplikace v detailu zobrazuje **všechny** položky balíčku – cizí jsou šedé se štítkem **jiný PM**. Co vidíte, to se přesně vyexportuje do importu.

- Cizí položky můžete nechat beze změny.
- Při exportu vás aplikace **upozorní**, pokud mezi upravenými balíčky jsou sdílené, ať si změnu zkoordinujete s druhým PM.

---

## 5. Uložení práce a návrat k rozpracovanému

### Automatické ukládání (stejný počítač)
Po každé změně se práce **sama ukládá** do prohlížeče – v hlavičce vidíte *„✓ uloženo HH:MM"*. Když appku zavřete a za pár dní ji na **stejném počítači a v stejném prohlížeči** zase otevřete, nahoře se objeví nabídka:

- **Pokračovat** – načte vaši rozpracovanou práci a jdete dál.
- **Zahodit a začít znovu** – smaže uloženou rozpracovanou práci.

### Přenosný soubor (jiný počítač / záloha)
Když budete chtít pokračovat **na jiném počítači** nebo si udělat zálohu:

1. V seznamu balíčků klikněte **💾 Uložit rozpracované** → stáhne se soubor `rozpracovane_balicky_<PM>_<datum>.json`. Uložte ho třeba na OneDrive.
2. Na druhém počítači otevřete aplikaci, načtěte `balicky_sk.xlsx`, a pak dejte **📂 Načíst rozpracované** a vyberte ten `.json`. (Načíst rozpracované jde i hned na úvodní obrazovce.)

> Tip: Automatické ukládání je vázané na konkrétní prohlížeč/počítač. Pokud měníte počítače nebo chcete jistotu, používejte přenosný soubor.

---

## 6. Mazání položek a celých balíčků

- **Položku** smažete tlačítkem *Smazat* v detailu balíčku (lze *Vrátit*).
- **Celý balíček** smažete ikonou **🗑** v seznamu. Přesune se do přehledu **🗑 Smazané balíčky**.

Smazané balíčky se **nedávají** do běžného ERP importu – jsou určené k ručnímu zneaktivnění/smazání v ERP. V přehledu *Smazané balíčky* je můžete **Obnovit** nebo vyexportovat jejich **Seznam (XLSX/CSV)** se sloupcem `AKCE_V_ERP`.

---

## 7. Export do ERP

Když máte balíčky hotové, v seznamu balíčků použijte řádek **„ERP import upravených balíčků"**:

- **📤 ERP import (ZIP po balíčcích)** – ZIP, kde je:
  - **jeden importní soubor `.txt` pro každý** upravený balíček (název = pořadí + název balíčku),
  - **`_SEZNAM_balicku.csv`** – soupis balíčků (pořadí, soubor, ID balíčku, počet položek, sdílení, PM).
- **📄 Přehled (XLSX)** – tabulka všech vašich upravených balíčků jen pro **vaši kontrolu/archiv** (není to importní soubor pro ERP).

Každý importní `.txt` je přesně ve formátu, který ERP očekává: **oddělený tabulátory, kódování Windows-1250, sloupce `ID` a `POPIS`**, jeden soubor = kompletní seznam položek balíčku. (Ověřeno: shoduje se bajt po bajtu se vzorovým souborem z ERP.)

V detailu balíčku je navíc tlačítko **📤 ERP soubor balíčku** – stáhne `.txt` jen pro právě otevřený balíček (rychlý import jednoho balíčku).

Do exportu jdou **jen balíčky, které jste změnil vy** (váš PM). Smazané balíčky tam nejsou (ty řešíte ručně podle jejich seznamu).

### Uložit změny (aktualizovaný xlsx + log)
Tlačítko **💾 Uložit změny** v hlavičce stáhne `balicky_sk_aktualizovany_<datum>.xlsx` – aktualizovaný stav balíčků + list **`log_zmen`** se všemi provedenými záměnami, přidáními a mazáními (kdo, kdy, co za co).

---

## 8. Typické scénáře

**A) Připravím balíčky dnes, import udělám za dva dny (stejný PC)**
1. Načtu balíčky, vyberu PM, upravím staré položky.
2. Nic dalšího nedělám – práce se ukládá sama.
3. Za dva dny otevřu aplikaci → **Pokračovat** → dokončím → **ERP import**.

**B) Začnu v práci, dokončím doma (jiný PC)**
1. V práci upravím, co stihnu, dám **💾 Uložit rozpracované** → uložím `.json` na OneDrive.
2. Doma otevřu aplikaci, načtu `balicky_sk.xlsx`, dám **📂 Načíst rozpracované**, vyberu `.json` → pokračuji.

**C) Dokončím a předám do ERP**
1. Projdu všechny balíčky (cíl: u každého *hotovo*).
2. **📤 ERP import (ZIP po balíčcích)** = importní `.txt` soubory + soupis; k tomu **📄 Seznam smazaných balíčků** k ručnímu zneaktivnění v ERP.
3. Pro vlastní archiv případně **💾 Uložit změny** (aktualizovaný xlsx + log) nebo **📄 Přehled (XLSX)**.

---

## 9. Řešení potíží

- **„Tento soubor je používán"** při načítání balíčků → `balicky_sk.xlsx` je otevřený v Excelu. Zavřete ho a načtěte znovu.
- **Nenačítají se návrhy / hláška, že katalog se nenačetl** → `katalog.js` musí ležet ve stejné složce jako `aktualizace_balicku.html`.
- **Po otevření appky nevidím rozpracovanou práci** → buď jste na jiném počítači/prohlížeči (použijte přenosný `.json`), nebo byla práce *Zahozena*.
- **Špatný návrh náhrady** → použijte v okně *Zaměnit* fulltextové hledání; případně nahlaste příklad, doladí se pravidla.

---

## 10. Důležité zásady

- Vstupní soubory (`balicky_sk.xlsx`, `katalog.js`) se **nikdy nepřepisují**.
- Do ERP importu jdou **jen vaše** upravené balíčky; sdílené obsahují i položky jiných PM (ERP nahrazuje celý balíček) – koordinujte.
- Smazané balíčky řešíte v ERP **ručně** podle jejich seznamu.
- Pro jistotu napříč počítači používejte **přenosný soubor** rozpracované práce.
