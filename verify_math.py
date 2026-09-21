import openpyxl

wb = openpyxl.load_workbook('Cuentas Fiestas Patrias.xlsx')
s1 = wb['Compras']
s2 = wb['Consumo por Integrante']
s3 = wb['División de Cuentas']

tot_c1 = 122251
tot_c2 = 61141
tot_c3 = 11790
tot_c4 = 5570
gran_total = tot_c1 + tot_c2 + tot_c3 + tot_c4

c1 = {
    'Chorizo': 14670,
    'Lomo Liso': 21226,
    'Lomo vetado': 19560,
    'Costillar Chileno': 19238,
    'Costillar cerdo': 14957,
    'Trutro': 0,
    'Pechuga': 0,
    'Carbón': 0,
    'Pan Choripán': 0,
    'Trigo Mote (1kg)': 2490,
    'Huesillos (1kg)': 10990,
    'Chancaca y Azúcar': 0,
    'Restante': 19120
}

c2 = {
    'Chorizo': 4890,
    'Lomo Liso': 0,
    'Lomo vetado': 33693,
    'Costillar Chileno': 0,
    'Costillar cerdo': 0,
    'Trutro': 16353,
    'Pechuga': 6205,
    'Carbón': 0,
    'Pan Choripán': 0,
    'Trigo Mote (1kg)': 0,
    'Huesillos (1kg)': 0,
    'Chancaca y Azúcar': 0,
    'Restante': 0
}

c3 = {
    'Chorizo': 0,
    'Lomo Liso': 0,
    'Lomo vetado': 0,
    'Costillar Chileno': 0,
    'Costillar cerdo': 0,
    'Trutro': 0,
    'Pechuga': 0,
    'Carbón': 8400,
    'Pan Choripán': 0,
    'Trigo Mote (1kg)': 0,
    'Huesillos (1kg)': 0,
    'Chancaca y Azúcar': 3390,
    'Restante': 0
}

c4 = {
    'Chorizo': 0,
    'Lomo Liso': 0,
    'Lomo vetado': 0,
    'Costillar Chileno': 0,
    'Costillar cerdo': 0,
    'Trutro': 0,
    'Pechuga': 0,
    'Carbón': 0,
    'Pan Choripán': 5570,
    'Trigo Mote (1kg)': 0,
    'Huesillos (1kg)': 0,
    'Chancaca y Azúcar': 0,
    'Restante': 0
}

tot_items = {}
for k in c1:
    tot_items[k] = c1[k] + c2[k] + c3[k] + c4[k]

integrantes = [
    'Pamela', 'Cote', 'Joaquín', 'Ignacio', 'Kena',
    'Ale', 'Mauri', 'Mindy', 'Monse', 'Gustavo',
    'Coni', 'Miriam', 'Abuelo Coni', 'Mamá Mauri', 'Papá Mauri'
]

consumo = {}
for m in integrantes:
    if m == 'Monse':
        consumo[m] = ['Trigo Mote (1kg)', 'Huesillos (1kg)', 'Chancaca y Azúcar', 'Restante']
    elif m == 'Miriam':
        consumo[m] = [
            'Costillar Chileno', 'Costillar cerdo', 'Trutro', 'Pechuga', 
            'Carbón', 'Trigo Mote (1kg)', 'Huesillos (1kg)', 'Chancaca y Azúcar', 'Restante'
        ]
    else:
        consumo[m] = list(tot_items.keys())

conteos = {item: sum(1 for m in integrantes if item in consumo[m]) for item in tot_items}
valor_persona = {item: tot_items[item] / conteos[item] for item in tot_items}

print("=== DESGLOSE DE ITEMS ===")
for item, total in tot_items.items():
    print(f"{item:20}: Total=${total:6d} | Personas={conteos[item]:2d} | Cuota individual=${valor_persona[item]:8.2f}")

print("\n=== CUOTAS Y LIQUIDACIÓN POR INTEGRANTE ===")
cuotas_base = {}
for m in integrantes:
    cuotas_base[m] = sum(valor_persona[it] for it in consumo[m])

cuota_kena_exacta = cuotas_base['Kena']
cuota_kena_split = cuota_kena_exacta / 4
sponsors_kena = ['Pamela', 'Ale', 'Mindy', 'Joaquín']
aportes = {'Ignacio': 183392, 'Mindy': 11790, 'Joaquín': 5570}

print(f"Cuota de Kena: ${cuota_kena_exacta:.2f} -> $ {cuota_kena_split:.2f} c/u para Pamela, Ale, Mindy y Joaquín\n")

for m in integrantes:
    base = cuotas_base[m]
    extra_kena = cuota_kena_split if m in sponsors_kena else 0
    tot_pagar = 0 if m == 'Kena' else (base + extra_kena)
    pagado = aportes.get(m, 0)
    saldo_transferir = 0 if (m == 'Ignacio' or m == 'Kena') else max(0, round(tot_pagar) - pagado)
    
    nota = ""
    if m in sponsors_kena:
        nota = f"(Base ${round(base):,d} + Kena ${round(extra_kena):,d})"
    elif m == 'Kena':
        nota = "(Cubierto por familiares)"
        
    print(f"{m:15}: Cuota Total=${round(tot_pagar):6,d} {nota:32} | Aporte Compras=${pagado:7,d} | Saldo a Transferir=${saldo_transferir:6,d}")

print("\n=== TOTALES ===")
print(f"Suma exacta cuotas: ${sum(cuotas_base.values()):.2f}")
print(f"Gran total compras: ${gran_total:,}")



