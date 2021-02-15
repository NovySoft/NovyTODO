# NovyTODO

Egy olcsó program ami átírja a házifeladatokat a teamsből és a krétából a microsoft TODO-ba.
A kódot ne próbáld megérteni, mert én sem tudom hogyan működik és néha miért nem :)

## Setup
Hozz létre egy `config.json` fájlt a `config.example.json` alapján.

És persze szükség van egy microsoft azure-os alkalmazásra, ami beírja a teendőket

## Kérdések

- Miért pont puppeteer?

Mert a teams apihoz való hozzáféréshez adminisztrátor jogok kellenek

- Miért ilyen borzalmas a kód?

Mert nekem csak az a lényeg, hogy átimportálja a teendőket egy helyre és ne kavarodjak bele a határidőkbe.

De ha neked van ötleted a tovább fejlesztésre esetleg kód optimalizálásra akkor pull requestedet szívesen várjuk!

Van néhány ismert problémája a kódnak, melyek random szoktak jelentkezni, de a legsúlyosabb jelenleg a `Navigator Timeout` problémák
