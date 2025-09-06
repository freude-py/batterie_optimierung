import openpyxl
import dateparser
import matplotlib.pyplot as plt
import csv
import pulp
import requests
from datetime import datetime

def lade_daten(excel_datei):
    wb = openpyxl.load_workbook(excel_datei, data_only=True)
    ws = wb.active
    daten = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        datum = row[0]
        try:
            preis = float(str(row[1]).replace(',', '.'))
        except (ValueError, TypeError):
            preis = 0
        try:
            verbrauch = float(str(row[5]).replace(',', '.'))
        except (ValueError, TypeError):
            verbrauch = 0
        daten.append([datum, preis, verbrauch])
    return daten

def filter_zeitraum(daten, von, bis):
    von_dt = dateparser.parse(von)
    bis_dt = dateparser.parse(bis)
    return [row for row in daten if dateparser.parse(str(row[0])) and von_dt <= dateparser.parse(str(row[0])) <= bis_dt]

def batterie_optimierung(daten, batterie_kapazitaet=10.0, lade_menge=0.5, preisdiff=5):
    batterie = 0.0
    lade_preise = []
    aktionen = []
    for i, (datum, preis, verbrauch) in enumerate(daten):
        if batterie < batterie_kapazitaet and (not lade_preise or preis <= min([d[1] for d in daten[max(0, i-96):i+1]], default=preis)):
            batterie += lade_menge
            lade_preise.append(preis)
            betrag = lade_menge * preis
            aktionen.append((datum, 'Laden', preis, batterie, None, betrag))
        elif batterie >= lade_menge and lade_preise and preis - min(lade_preise) >= preisdiff and verbrauch > 0:
            batterie -= lade_menge
            ersparnis = preis - min(lade_preise)
            betrag = -lade_menge * preis
            aktionen.append((datum, 'Entladen', preis, batterie, ersparnis, betrag))
            lade_preise.remove(min(lade_preise))
    return aktionen

def batterie_optimierung_neu(daten, batterie_kapazitaet=10.0, lade_menge=0.5, ladung_min=10, ladung_max=90, speicher_verluste=10, entlade_verluste=10):
    import pulp

    speicher_verluste = speicher_verluste / 100.0  # Prozent zu Anteil
    entlade_verluste = entlade_verluste / 100.0

    n = len(daten)

    # Umrechnung Prozent auf Anteil
    ladung_min_anteil = ladung_min / 100.0
    ladung_max_anteil = ladung_max / 100.0

    # Variablen
    l = [pulp.LpVariable(f"l_{t}", lowBound=0, upBound=lade_menge) for t in range(n)]
    # e[t] ist die tatsächlich genutzte Energie (Verbrauch)
    e = [pulp.LpVariable(f"e_{t}", lowBound=0, upBound=daten[t][2]) for t in range(n)]
    s = [pulp.LpVariable(f"s_{t}", lowBound=ladung_min_anteil * batterie_kapazitaet, upBound=ladung_max_anteil * batterie_kapazitaet) for t in range(n+1)]

    prob = pulp.LpProblem("BatterieOptimierung", pulp.LpMaximize)

    # Zielfunktion: Erlöse - Kosten (mit Verlusten)
    prob += pulp.lpSum([
        daten[t][1] * e[t] - daten[t][1] * l[t] / (1 - speicher_verluste)
        for t in range(n)
    ])

    # Anfangs- und Endfüllstand
    prob += s[0] == ladung_min_anteil * batterie_kapazitaet
    prob += s[n] == ladung_min_anteil * batterie_kapazitaet

    # Speicherentwicklung
    for t in range(n):
        # Für den Verbrauch e[t] muss die Batterie um e[t] / (1 - entlade_verluste) entladen werden
        prob += s[t+1] == s[t] + l[t] - e[t] / (1 - entlade_verluste)

    # Batterie am Tagesende auf ladung_min
    last_tag = None
    for t in range(n):
        datum = str(daten[t][0])
        tag = datum[:10]
        if last_tag and tag != last_tag:
            prob += s[t] == ladung_min_anteil * batterie_kapazitaet
        last_tag = tag

    # Lösen
    prob.solve(pulp.PULP_CBC_CMD(msg=0))

    # Ergebnisse aufbereiten
    aktionen = []
    for t in range(n):
        datum, preis, verbrauch = daten[t]
        lade_val = round(l[t].value() or 0, 4)  # Menge, die in die Batterie kommt
        verbrauch_val = round(e[t].value() or 0, 4)  # Tatsächlich genutzte Energie
        entlade_val = round(verbrauch_val / (1 - entlade_verluste), 4) if verbrauch_val > 0 else 0.0  # Entnommene Energie aus Batterie
        speicher_val = round(s[t+1].value() or 0, 4)
        menge_i = round(verbrauch, 4)
        preis = round(preis, 4)
        netz_laden = round(lade_val / (1 - speicher_verluste), 4) if lade_val > 0 else 0.0
        if lade_val > 0:
            betrag = round(netz_laden * preis, 4)
            aktionen.append((datum, "Laden", preis, speicher_val, None, betrag, menge_i, netz_laden))
        elif verbrauch_val > 0:
            betrag = round(-verbrauch_val * preis, 4)
            aktionen.append((datum, "Entladen", preis, speicher_val, None, betrag, menge_i, entlade_val))
        else:
            aktionen.append((datum, "Nichts", preis, speicher_val, None, 0.0, menge_i, 0.0))
    return aktionen

def schreibe_csv(aktionen, dateiname):
    with open(dateiname, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f, delimiter=';')
        writer.writerow(['Datum', 'Aktion', 'Menge_V', 'Menge_a', 'Preis', 'Fuellstand', 'Ersparnis', 'Betrag'])
        def excel_float(val, digits):
            return (f"{val:.{digits}f}".replace(".", ",") if isinstance(val, float) else "")
        last_day = None
        last_month = None
        sum_menge_v = sum_menge_a = sum_preis = sum_fuellstand = sum_betrag = 0.0
        sum_menge_v_monat = sum_menge_a_monat = sum_preis_monat = sum_fuellstand_monat = sum_betrag_monat = 0.0
        sum_menge_v_gesamt = sum_menge_a_gesamt = sum_preis_gesamt = sum_fuellstand_gesamt = sum_betrag_gesamt = 0.0
        for aktion in aktionen:
            preis = float(aktion[2]) if aktion[2] not in (None, "") else 0.0
            fuellstand = float(aktion[3]) if aktion[3] not in (None, "") else 0.0
            ersparnis = float(aktion[4]) if aktion[4] not in (None, "") else 0.0
            betrag = float(aktion[5]) if aktion[5] not in (None, "") else 0.0
            menge_i = float(aktion[6]) if aktion[6] not in (None, "") else 0.0
            menge_a = float(aktion[7]) if aktion[7] not in (None, "") else 0.0
            datum = str(aktion[0])
            tag = datum[:10]
            monat = datum[:7]
            # Summen sammeln
            sum_menge_v += menge_i
            sum_menge_a += menge_a
            sum_preis += preis
            sum_fuellstand += fuellstand
            sum_betrag += betrag
            sum_menge_v_monat += menge_i
            sum_menge_a_monat += menge_a
            sum_preis_monat += preis
            sum_fuellstand_monat += fuellstand
            sum_betrag_monat += betrag
            sum_menge_v_gesamt += menge_i
            sum_menge_a_gesamt += menge_a
            sum_preis_gesamt += preis
            sum_fuellstand_gesamt += fuellstand
            sum_betrag_gesamt += betrag
            writer.writerow([
                datum,
                aktion[1],
                excel_float(menge_i, 4),
                excel_float(menge_a, 4),
                excel_float(preis, 4),
                excel_float(fuellstand, 4),
                excel_float(ersparnis, 4),
                excel_float(betrag, 4)
            ])
            # Summenzeile am Tagesende
            if last_day and tag != last_day:
                writer.writerow([
                    last_day,
                    "Tagessumme",
                    excel_float(sum_menge_v, 4),
                    excel_float(sum_menge_a, 4),
                    excel_float(sum_preis, 4),
                    excel_float(sum_fuellstand, 4),
                    "",
                    excel_float(sum_betrag, 4)
                ])
                sum_menge_v = sum_menge_a = sum_preis = sum_fuellstand = sum_betrag = 0.0
            # Summenzeile am Monatsende
            if last_month and monat != last_month:
                writer.writerow([
                    last_month,
                    "Monatssumme",
                    excel_float(sum_menge_v_monat, 4),
                    excel_float(sum_menge_a_monat, 4),
                    excel_float(sum_preis_monat, 4),
                    excel_float(sum_fuellstand_monat, 4),
                    "",
                    excel_float(sum_betrag_monat, 4)
                ])
                sum_menge_v_monat = sum_menge_a_monat = sum_preis_monat = sum_fuellstand_monat = sum_betrag_monat = 0.0
            last_day = tag
            last_month = monat
        # Summenzeile für den letzten Tag
        if last_day:
            writer.writerow([
                last_day,
                "Tagessumme",
                excel_float(sum_menge_v, 4),
                excel_float(sum_menge_a, 4),
                excel_float(sum_preis, 4),
                excel_float(sum_fuellstand, 4),
                "",
                excel_float(sum_betrag, 4)
            ])
        # Summenzeile für den letzten Monat
        if last_month:
            writer.writerow([
                last_month,
                "Monatssumme",
                excel_float(sum_menge_v_monat, 4),
                excel_float(sum_menge_a_monat, 4),
                excel_float(sum_preis_monat, 4),
                excel_float(sum_fuellstand_monat, 4),
                "",
                excel_float(sum_betrag_monat, 4)
            ])
        # Gesamtsumme
        writer.writerow([
            "Gesamt",
            "Gesamtsumme",
            excel_float(sum_menge_v_gesamt, 4),
            excel_float(sum_menge_a_gesamt, 4),
            excel_float(sum_preis_gesamt, 4),
            excel_float(sum_fuellstand_gesamt, 4),
            "",
            excel_float(sum_betrag_gesamt, 4)
        ])

def visualisiere(aktionen):
    lade_zeitpunkte = [dateparser.parse(str(a[0])) for a in aktionen if a[1] == 'Laden']
    lade_preise_plot = [a[2] for a in aktionen if a[1] == 'Laden']
    entlade_zeitpunkte = [dateparser.parse(str(a[0])) for a in aktionen if a[1] == 'Entladen']
    entlade_preise_plot = [a[2] for a in aktionen if a[1] == 'Entladen']
    plt.figure(figsize=(14, 6))
    plt.plot(lade_zeitpunkte, lade_preise_plot, 'go', label='Laden')
    plt.plot(entlade_zeitpunkte, entlade_preise_plot, 'ro', label='Entladen')
    plt.xlabel('Zeit')
    plt.ylabel('Preis (Cent/kWh)')
    plt.title('Batterie-Lade- und Entladezeitpunkte')
    plt.legend()
    plt.grid(True)
    plt.tight_layout()
    plt.show()

def berechne_summen(aktionen, lade_menge):
    ladekosten = sum(a[2] * lade_menge for a in aktionen if a[1] == 'Laden')
    entladeerlöse = sum(a[2] * lade_menge for a in aktionen if a[1] == 'Entladen')
    differenz = entladeerlöse - ladekosten
    return ladekosten, entladeerlöse, differenz

def marktpreise_holen():
    url = "https://api.awattar.at/v1/marketdata"
    try:
        response = requests.get(url)
        print("Awattar API Response Status:", response.status_code)
        if response.status_code == 200:
            data = response.json().get("data", [])
            print("Awattar Zeitfenster:")
            for entry in data:
                start_dt = datetime.fromtimestamp(entry["start_timestamp"] / 1000)
                end_dt = datetime.fromtimestamp(entry["end_timestamp"] / 1000)
                print(f"{start_dt.strftime('%Y-%m-%d %H:%M')} - {end_dt.strftime('%Y-%m-%d %H:%M')}: {entry['marketprice']} {entry['unit']}")
        else:
            print("Awattar API Response:")
            print(response.text[:1000])
    except Exception as e:
        print("Fehler beim Abrufen der Marktpreise:", e)

def marktpreise_von_entsoe():
    import requests
    from datetime import datetime, timedelta

    token = "1d8d5cf8-299c-4b8f-b142-3f3b294f2b77"
    morgen = datetime.now() + timedelta(days=1)
    start = morgen.strftime("%Y%m%d0000")
    end = morgen.strftime("%Y%m%d2300")

    #https://web-api.transparency.entsoe.eu/api?securityToken=1d8d5cf8-299c-4b8f-b142-3f3b294f2b77&documentType=A44&in_Domain=10YAT-APG------L&out_Domain=10YAT-APG------L&periodStart=202509020000&periodEnd=202509022300


    url = (
    f"https://web-api.transparency.entsoe.eu/api?securityToken={token}"
    f"&documentType=A44"
    f"&in_Domain=10YAT-APG------L"
    f"&out_Domain=10YAT-APG------L"
    f"&periodStart={start}"
    f"&periodEnd={end}"
)
    headers = {"User-Agent": "Mozilla/5.0"}
    try:
        response = requests.get(url, headers=headers)
        print("ENTSO-E API Response Status:", response.status_code)
        if response.status_code == 200:
            import xml.etree.ElementTree as ET
            root = ET.fromstring(response.content)
            ns = {'ns': 'urn:iec62325.351:tc57wg16:451-3:publicationdocument:7:0'}
            prices = []
            for time_series in root.findall('.//ns:TimeSeries', ns):
                for period in time_series.findall('.//ns:Period', ns):
                    start_time = period.find('ns:timeInterval/ns:start', ns).text
                    for point in period.findall('ns:Point', ns):
                        position = int(point.find('ns:position', ns).text)
                        price = float(point.find('ns:price.amount', ns).text)
                        dt = datetime.strptime(start_time, "%Y-%m-%dT%H:%MZ") + timedelta(hours=position-1)
                        prices.append((dt, price))
            print("ENTSO-E Day-Ahead Preise für Österreich (EUR/MWh):")
            for dt, price in prices:
                print(f"{dt.strftime('%Y-%m-%d %H:%M')}: {price:.2f} EUR/MWh")
        else:
            print("ENTSO-E API Response:")
            print(response.text[:1000])
    except Exception as e:
        print("Fehler beim Abrufen der ENTSO-E Marktpreise:", e)

def main():

    excel_datei_input = 'h0_preise_15minuten.xlsx'
    print("\nRentabilitätsrechnung für eine Batterie mit 15 Minmuten Preisen und 15 Minuten Verbräuchen")
    print("==========================================================================================")
    von_default = "2026-01-01"
    bis_default = "2026-01-03"
    von = input(f"Von-Datum  [Default: {von_default}]: ")
    bis = input(f"Bis-Datum  [Default: {bis_default}]: ")
    if not von: von = von_default
    if not bis: bis = bis_default

    # Neue Eingaben mit Default
    batterie_kapazitaet_default = 10.0
    lade_menge_default = 0.5
    ladung_min_default = 10
    ladung_max_default = 90
    speicher_verluste_default = 10
    entlade_verluste_default = 10
    batterie_kapazitaet = input(f"Batterie-Kapazität in kWh       [Default: {batterie_kapazitaet_default}]: ")
    lade_menge = input(f"Lademenge pro 15 Minuten in kWh [Default:  {lade_menge_default}]: ")
    ladung_min = input(f"Minimaler Ladezustand in %      [Default: {ladung_min_default}  ]: ")
    ladung_max = input(f"Maximaler Ladezustand in %      [Default: {ladung_max_default}  ]: ")
    speicher_verluste = input(f"Speicherverluste in %           [Default: {speicher_verluste_default}  ]: ")
    entlade_verluste = input(f"Entladeverluste  in %           [Default: {entlade_verluste_default}  ]: ")
    batterie_kapazitaet = float(batterie_kapazitaet) if batterie_kapazitaet else batterie_kapazitaet_default
    lade_menge = float(lade_menge) if lade_menge else lade_menge_default
    ladung_min = float(ladung_min) if ladung_min else ladung_min_default
    ladung_max = float(ladung_max) if ladung_max else ladung_max_default
    speicher_verluste = float(speicher_verluste) if speicher_verluste else speicher_verluste_default
    entlade_verluste = float(entlade_verluste) if entlade_verluste else entlade_verluste_default

    # Setze Zeit auf Tagesanfang/ende
    von_dt = dateparser.parse(von + " 00:00")
    bis_dt = dateparser.parse(bis + " 23:45")
    daten = lade_daten(excel_datei_input)
    daten = [row for row in daten if dateparser.parse(str(row[0])) and von_dt <= dateparser.parse(str(row[0])) <= bis_dt]

    aktionen = batterie_optimierung_neu(daten, batterie_kapazitaet, lade_menge, ladung_min, ladung_max, speicher_verluste, entlade_verluste)
    schreibe_csv(aktionen, 'batterie_aktionen.csv')
    #visualisiere(aktionen)
    ladekosten, entladeerlöse, differenz = berechne_summen(aktionen, lade_menge)
    print(f"\nGesamte Ladekosten: {ladekosten:.2f} Cent")
    print(f"Gesamte Entladeerlöse: {entladeerlöse:.2f} Cent")
    print(f"Differenz (Ersparnis): {differenz:.2f} Cent")

if __name__ == "__main__":
    #marktpreise_holen()
    #marktpreise_von_entsoe()
    main()
    #    $ curl "https://api.awattar.at/v1/marketdata"      # Preise für die nächsten 24 Stunden
