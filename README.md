 

Súhrn požiadaviek  

1. potrebné výlepšenia alebo opravy nedostatkov 

Celková revízia a oprava kódu pri zachovaní dizajnu  súborov (Code.gs, CSS, Index.html, JavaScript.html) pre aplikáciu app script v goole sheet. 

 

Nefunkčnosť základnej synchronizácie medzi web aplikaciou a google sheetom. Funguje Edit, Duplikovanie záznamu a zmena v stĺpci  “Staus”. Opačne táto synchonizácia nefunguje ked zmením hodnoty priamo v google sheet. Po zmene v riadkoch príslušneho hárka v zaznámoch od riadku 17 by sa malo automaticky spustiť znovu načítanie datatabuliek vo web aplikácií.      

 

 

Nejde pridať záznam po kliknutí na tlačidlo “Pridaj záznam” malo by otvoriť modal okno kde sa dajú zapísať všetky položky aktuálnej karty. 

 

V stĺpci s  názvom “Označ” so zaškrtávacím tlačidlom treba doplniť aj možnosť označovanie alebo odznačovanie riadkov pomocou ctrl+click alebo shift+click  okrem označenia pomocou zaškrtávacieho tlačida. A to ak klikne vedľa tohto zaškrtávacieho tlačidla. 

 

Treba doplniť vyfarbovanie aj globálneho filtra momentálne sa nevyfarbí  do farby záložky ak je tam napísaná hodnota pre filtrovanie podobne ako sa vyfarbuju bunky stĺpcových filtrov. Globalný filter by sa mal vyfarbiť do takej istej farby ako farebná téma záložky ak je v tomto filtri nejaká hodnota. . Podobná funcia  ako sa vyfarbuje rozbaľovací panel filtrov ak sú nejaké použité. 

 

 

2. Typy polí v modal okne  pre pridaj uprav duplikuj alebo vo filtroch.   

-ak je v hárku v nazve hlavičky v riadku 16 “názov” a za ním hneď symbol”/*” takže v príklade “Názov1/*” tak ide o typ stĺpca, ktorý sa bude v datatabuľke v päte datatabuľky spočítavať v súčtovom riadku.  

 

-ak je v hárku v nazve hlavičky v riadku 16 “názov” a za ním hneď symbol”/-” takže v príklade “Názov1/-” tak ide o typ stĺpca, ktorý sa bude v modal okne aj vo filtroch zobrazovať ako bunka výberového typu. Rozsah pre výber tejto bunky bude vždy v hárku v riadku 15 hned nad týmto názvom stĺpca. 

 

-ak je v hárku v nazve hlavičky v riadku 16 “názov” a za ním hneď symbol”/att” takže v príklade “Názov1/att“ tak ide o typ stĺpca, ktorý sa bude slúžiť na nahranie prílohy v modal okne. 

syntax pre prílohy: Treba implementovať požiadavku na rozpoznávanie hlavičiek končiacich na "/att". 

Dynamický formulár: V modálnom okne sa pre stĺpce s príponou /att automaticky zobrazuje pole s názvom bez tejto prípony /att na nahratie súboru , ktoré na mobilných zariadeniach ponúka aj možnosť priameho fotenia. 

Automatické ukladanie na Google Drive: Nahrané súbory sa ukladajú do štruktúrovaných priečinkov na Disku v tvare: [Hlavný priečinok]/att_[Názov hárku]/[ID záznamu]/[názov súboru]. 

Odkaz v tabuľke: Do bunky v Google Sheets sa po nahratí súboru uloží priamy, klikateľný odkaz na daný súbor. 

 

-ak je v hárku v nazve hlavičky v riadku 16 “názov” a za ním hneď symbol”/emailto” takže v príklade “Názov1/emailto” tak ide o typ stĺpca, ktorý  bude obsahovať emailovú adresu ktorá bude slúžiť na odoslanie emailu na túto adresu cez funkciu  Email  tlačidla. Ale vyskakovacie okno musí ponúknuť aj možnosť  len zobraziť tento predvolený email a tiež možnosť zadať email alebo viaceré emaily do políčka odedelne ciarkov ako email1, email2, ...  ak sa do tohto políčka napíše email alebo viaceré emaily export sa pošle len na tieto emaily. Ak bude toto políčko prázdne export sa pošle na predvolený email z tabuľky.      

-treba skontrolovať potom funčnosť tohto tlačida Email v tejto verzii kódu vyskočilo okno ale stale načítavalo.  

 

 

3. ostatné nedostatky a požiadavky.   

 

-treba skontrolovať funčnosť exportov PDF, PDF s logom, HTML a XLS tak aby sa exportovali len riadky ktoré sú filtrované a stĺpce ktoré sú aktuálne zobrazené cez tlačidlo stĺpce. Ak sú niektoré riadky vyrané cez kliknutie alebo zaškrtávacie tlačidlo použi len tieto vybraté riadky do exportu. V obidvoch prípadoch sa musí do exportu zahrnúť aj riadok súčtu. 

 

-doplň systémový stĺpec pred stĺpec “id” s nazvom “QR” tento stĺpec bude slúžiť na zobrazenie QR kódu hodnoty id riadku. Kód sa musí zobraziť po kliknutí v riadku na nejakú ikonku QR kódu. 

-tlačidlo refresh a automatický refresh po 5 min nefunguje. Píše chybu Dáta sas nepodarilo obnoviť. Skontoluj funfčnosť tak aby sa po kliknutí alebo opočítaní nového času 2 minúty znovu načítali dáta datatabuľky na pozadí z príslušných hárkov.  

 

 

 

 

 

 

 
