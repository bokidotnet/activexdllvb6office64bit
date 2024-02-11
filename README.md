# activexdllvb6office64bit
Pomoću ove ActiveX Exe biblioteke se može izvršiti "ubacivanje" i korišćenje bilo kog ActiveX Dll napisanog u VB6 u pre svega MS Office 64-bit verzije VBA a da se ne menja već urađena ActiveX dll biblioteka.

Potrebno je ovaj ActiveX Exe dodati kao referencu u VBA projekat kao i svaki drugi ActiveX Dll kako bi se isti koristio pored ActiveX dll koji se koristi u VBA projektu.

U prilogu je primer za ActiveX Dll MUPLKLib.

Potrebno je registrovati MUPLKLib.dll sa regsvr32 odnosno kako je opisano u procitajme.txt.

CelikApi.dll je potrebno da bude vidljiv od strane operativnog sistema kako bi biblioteka MUPLKLib.dll mogla da pristupi CelikApi.

NAPOMENE:
Nije testirano sa poslednjom verzijom CelikApi.dll 1.3.3 koja verovatno neće biti podržana zbog izvršenih poslednjih promena od strane MUP RS.
Iz tog razloga ali i kada je napravljen ActiveX Dll MUPLKLib su date obe verzije CelikApi.dll sa kojima je razvijan tada ActiveX dll.
