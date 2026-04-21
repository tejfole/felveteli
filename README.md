# felveteli

## Ribbon (customUI) + Iktsz menü

A Ribbon a `workbook/felveteli_2026.xlsm` fájlba van beágyazva, a `customUI/customUI14.xml` részen.

### Jelenlegi Iktsz callback

- `btnIktsz` → `onAction="KitoltIktsz_TablaAutomatikusan"`

### Jelenlegi customUI callback lista (`customUI14.xml`)

- `KitoltIktsz_TablaAutomatikusan`
- `ToltsdIskolaAdatokatPirosSargaHibaval`
- `Publikalas_AzonositoPont_Sorszam_RangsorSzerint`
- `MasolasDiakadatbolRangsorba`
- `SzamoljEgyediPrioritasosRangsor`
- `TorolPonthatarA14`
- `SaveVersionedCopy`
- `GeneratePasswordsFromTableAndExportCSV_Final_UniqueWithLogClean`
- `ExportCSVFromActiveSheetTable_UniqueOktazon`
- `Import_Export_Into_ThisWorkbook_Diakadat`
- `Import_KozpontiFelveteli_Pontszamok`
- `SzamoljPontokatTombosen`
- `Ribbon_BizonyitvanyMatrix`
- `Ribbon_BizonyitvanyFrissites`
- `Ribbon_BizonyitvanyTeljes`
- `PBizonyitvany_UjratoltesEsTizedesBeallitas`
- `Ribbon_TeremNevsor_Refresh`
- `Ribbon_TeremNevsor_Generate`
- `Idopontok_Menu`

### Új Iktsz callbackek (wrapper belépési pontok)

- `Iktsz_Menu_Intezmenyi` → `lista.iktsz` kitöltés `isk_nev` csoportosítással
- `Iktsz_Menu_Hatarozat` → `lista.iktsz` egyedi, szekvenciális kitöltés (csak nem üres `hatarozat` sorokra)
- `Iktsz_Menu_Szobeli` → `diakadat.iktsz` feltételes, szekvenciális kitöltés (`bizottsag`, `datum_nap`, `mail`, `idopont_kiadva`)

A callbackek implementációja: `vba/munkalap/IktszMenuCallbacks.bas`.

### customUI frissítés (ha a bináris `.xlsm` nem módosul PR-ben)

1. Nyisd meg a `workbook/felveteli_2026.xlsm` fájlt Office RibbonX / Custom UI Editor eszközben.
2. Nyisd meg a `customUI/customUI14.xml` tartalmat.
3. A jelenlegi `btnIktsz` gombot cseréld erre:

```xml
<splitButton id="sbIktsz">
  <button id="btnIktszDefault"
          label="Iktsz (intézményi)"
          image="iktatosz"
          size="large"
          onAction="Iktsz_Menu_Intezmenyi"
          screentip="Intézményi iktsz"
          supertip="lista táblában iktsz kitöltése isk_nev csoportosítással." />
  <menu id="menuIktsz" label="Iktsz műveletek" itemSize="large">
    <button id="btnIktszIntezmenyi"
            label="Intézményi (isk_nev)"
            image="iktatosz"
            onAction="Iktsz_Menu_Intezmenyi" />
    <button id="btnIktszHatarozat"
            label="Határozat (egyedi)"
            imageMso="FileProperties"
            onAction="Iktsz_Menu_Hatarozat" />
    <button id="btnIktszSzobeli"
            label="Szóbeli (feltételes)"
            imageMso="AppointmentColorDialog"
            onAction="Iktsz_Menu_Szobeli" />
  </menu>
</splitButton>
```

4. Mentsd a customUI módosítást.
5. Excel VBA-ban importáld/frissítsd a `vba/munkalap/IktszMenuCallbacks.bas` modult.
6. Nyisd meg újra a munkafüzetet, és teszteld a 3 Iktsz menüpontot.
