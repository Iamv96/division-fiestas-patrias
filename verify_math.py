import openpyxl

wb = openpyxl.load_workbook('Cuentas Fiestas Patrias.xlsx')
s1 = wb['Compras']
s2 = wb['Consumo por Integrante']
s3 = wb['División de Cuentas']

tot_c1 = 122251
tot_c2 = 61141
tot_c3 = 11790
tot_c4 = 5570
tot_c5 = 46250
gran_total = tot_c1 + tot_c2 + tot_c3 + tot_c4 + tot_c5

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
    'Restante': 19120,
    'Pino (1a Empanada)': 0,
    'Pino (2a Empanada)': 0,
    'Mechada': 0,
    'Camarón Queso': 0,
    'Queso': 0,
    'Napolitana': 0
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
    'Restante': 0,
    'Pino (1a Empanada)': 0,
    'Pino (2a Empanada)': 0,
    'Mechada': 0,
    'Camarón Queso': 0,
    'Queso': 0,
    'Napolitana': 0
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
    'Restante': 0,
    'Pino (1a Empanada)': 0,
    'Pino (2a Empanada)': 0,
    'Mechada': 0,
    'Camarón Queso': 0,
    'Queso': 0,
    'Napolitana': 0
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
    'Restante': 0,
    'Pino (1a Empanada)': 0,
    'Pino (2a Empanada)': 0,
    'Mechada': 0,
    'Camarón Queso': 0,
    'Queso': 0,
    'Napolitana': 0
}

c5 = {
    'Chorizo': 0,
    'Lomo Liso': 0,
    'Lomo vetado': 0,
    'Costillar Chileno': 0,
    'Costillar cerdo': 0,
    'Trutro': 0,
    'Pechuga': 0,
    'Carbón': 0,
    'Pan Choripán': 0,
    'Trigo Mote (1kg)': 0,
    'Huesillos (1kg)': 0,
    'Chancaca y Azúcar': 0,
    'Restante': 0,
    'Pino (1a Empanada)': 17500,
    'Pino (2a Empanada)': 15000,
    'Mechada': 6000,
    'Camarón Queso': 3000,
    'Queso': 2250,
    'Napolitana': 2500
}

tot_items = {}
for k in c1:
    tot_items[k] = c1[k] + c2[k] + c3[k] + c4[k] + c5[k]

integrantes = [
    'Pamela', 'Cote', 'Joaquín', 'Ignacio', 'Kena',
    'Ale', 'Mauri', 'Mindy', 'Monse', 'Gustavo',
    'Coni', 'Miriam', 'Abuelo Coni', 'Mamá Mauri', 'Papá Mauri',
    'Franklin'
]

asado_items = [
    'Chorizo', 'Lomo Liso', 'Lomo vetado', 'Costillar Chileno', 'Costillar cerdo',
    'Trutro', 'Pechuga', 'Carbón', 'Pan Choripán', 'Trigo Mote (1kg)', 'Huesillos (1kg)',
    'Chancaca y Azúcar', 'Restante'
]

consumo = {}
for m in integrantes:
    consumo[m] = []
    # Asado:
    if m == 'Monse':
        consumo[m] += ['Trigo Mote (1kg)', 'Huesillos (1kg)', 'Chancaca y Azúcar', 'Restante']
    elif m == 'Miriam':
        consumo[m] += [
            'Costillar Chileno', 'Costillar cerdo', 'Trutro', 'Pechuga', 
            'Carbón', 'Trigo Mote (1kg)', 'Huesillos (1kg)', 'Chancaca y Azúcar', 'Restante'
        ]
    else:
        consumo[m] += list(asado_items)
    
    # Empanadas:
    if m in ['Joaquín', 'Ignacio', 'Gustavo', 'Mauri', 'Pamela', 'Kena']:
        consumo[m] += ['Pino (1a Empanada)', 'Pino (2a Empanada)']
    elif m == 'Mindy':
        consumo[m] += ['Pino (1a Empanada)']
    elif m == 'Ale':
        consumo[m] += ['Mechada']
    elif m == 'Franklin':
        consumo[m] += ['Mechada', 'Camarón Queso']
    elif m == 'Cote':
        consumo[m] += ['Queso', 'Napolitana']

conteos = {item: sum(1 for m in integrantes if item in consumo[m]) for item in tot_items}
valor_persona = {item: tot_items[item] / conteos[item] for item in tot_items}

print("=== DESGLOSE DE ITEMS ===")
for item, total in tot_items.items():
    print(f"{item:22}: Total=${total:6d} | Personas={conteos[item]:2d} | Cuota individual=${valor_persona[item]:8.2f}")

print("\n=== CUOTAS Y LIQUIDACIÓN POR INTEGRANTE ===")
cuotas_base = {}
for m in integrantes:
    cuotas_base[m] = sum(valor_persona[it] for it in consumo[m])

cuota_kena_exacta = cuotas_base['Kena']
cuota_kena_split = cuota_kena_exacta / 4
sponsors_kena = ['Pamela', 'Ale', 'Mindy', 'Joaquín']
aportes = {'Ignacio': 183392, 'Mindy': 11790, 'Joaquín': 51820}

print(f"Cuota de Kena: ${cuota_kena_exacta:.2f} -> ${cuota_kena_split:.2f} c/u para Pamela, Ale, Mindy y Joaquín\n")

total_transferencias_a_ignacio = 0
for m in integrantes:
    base = cuotas_base[m]
    extra_kena = cuota_kena_split if m in sponsors_kena else 0
    tot_pagar = 0 if m == 'Kena' else (base + extra_kena)
    pagado = aportes.get(m, 0)
    saldo_neto = round(tot_pagar) - pagado
    
    nota = ""
    if m in sponsors_kena:
        nota = f"(Base ${round(base):,d} + Kena ${round(extra_kena):,d})"
    elif m == 'Kena':
        nota = "(Cubierto por familiares)"
    
    if m == 'Ignacio':
        accion = "Organizador (Recauda)"
    elif m == 'Kena':
        accion = "Cubierto"
    elif m == 'Joaquín':
        accion = f"Saldo a favor: Ignacio le transfiere ${-saldo_neto:,d}"
    else:
        accion = f"Transfiere ${saldo_neto:,d} a Ignacio"
        total_transferencias_a_ignacio += saldo_neto
        
    print(f"{m:15}: Cuota Total=${round(tot_pagar):6,d} {nota:32} | Aporte Compras=${pagado:7,d} | {accion}")

print("\n=== TOTALES Y BALANCE DE TESORERÍA (IGNACIO) ===")
print(f"Suma exacta cuotas: ${sum(cuotas_base.values()):.2f}")
print(f"Gran total compras: ${gran_total:,}")
print(f"Total recaudado por Ignacio (13 transferencias): ${total_transferencias_a_ignacio:,}")
recuperacion_ignacio = aportes['Ignacio'] - round(cuotas_base['Ignacio'])
print(f"Monto que recupera Ignacio para quedar en cero de sus compras: ${recuperacion_ignacio:,}")
excedente_ignacio = total_transferencias_a_ignacio - recuperacion_ignacio
print(f"Excedente que le sobra a Ignacio en cuenta: ${excedente_ignacio:,}")
print(f"Monto a transferir a Joaquín: ${-(round(cuotas_base['Joaquín'] + cuota_kena_split) - aportes['Joaquín']):,}")



