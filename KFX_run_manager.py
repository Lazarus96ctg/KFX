import pandas as pd
import os
import re
import sys

# Caminho para o Excel
excel_path = r'C:\Users\nnavarrosimancas\Documents\KFX_scenario_manager\Controle de simulação_teste_automatização_rev1.xlsx'

# Diretório de saída
output_dir = os.path.join(os.path.dirname(excel_path), "scn_output")
os.makedirs(output_dir, exist_ok=True)

# Função para formatação de coordenadas
def format_xyz(value):
    if pd.isna(value) or value in ['', 'nan', 'NaN', 'NaT']:
        return "(0,0,0)"
    try:
        if isinstance(value, str) and re.match(r'^\([^)]+\)$', value):
            return value
        value_str = str(value)
        if value_str.startswith('(') and value_str.endswith(')'):
            value_str = value_str[1:-1]
        cleaned = re.sub(r'[^0-9,\-+.]+', ' ', value_str)
        parts = re.split(r'[, ]+', cleaned)
        nums = []
        for part in parts:
            if part.strip() and part.replace('.', '').replace('-', '').strip():
                part = part.replace(',', '.')
                nums.append(float(part))
        while len(nums) < 3:
            nums.append(0.0)
        return f"({nums[0]},{nums[1]},{nums[2]})"
    except Exception:
        return "(0,0,0)"

# Função para formatar caminhos
def format_path(path):
    if pd.isna(path) or str(path).strip() in ['', 'nan']:
        return "''"
    path_str = str(path).replace('/', '\\')
    return f"'{path_str}'"

try:
    # LER TODAS AS COLUNAS E LINHAS RELEVANTES SEM ALTERAR NOMES
    df = pd.read_excel(
        excel_path,
        header=0,  # Primeira linha é cabeçalho
        nrows=18    # Ajuste conforme necessário
    )
    df.columns = df.columns.str.strip()
    print(f"Planilha carregada com {len(df)} linhas e {len(df.columns)} colunas")
    print("Colunas encontradas:", df.columns.tolist())
    print("Colunas encontradas (repr):", [repr(col) for col in df.columns])

    # Preload all composition sheets
    comp_sheets = {
        'M01.JF1': pd.read_excel(excel_path, sheet_name='M01.JF1_comp'),
        'M01.JF3': pd.read_excel(excel_path, sheet_name='M01.JF3_comp'),
        'M01.JF4': pd.read_excel(excel_path, sheet_name='M01.JF4_comp'),
        'M07.JF1': pd.read_excel(excel_path, sheet_name='M07.JF1_comp'),
    }
    for key, sheet in comp_sheets.items():
        sheet.columns = sheet.columns.str.strip()

    # Valores fixos para colunas não presentes na planilha
    TMAX = "1e+030"
    MSTEP = 900000
    RES_TEMP = 72
    JET_PRESSURE = 39.42
    LIQUID_COMPOSITION = "(0,0,0,0,0,0,0,0,0,0,0,0)"
    GAS_CLOUD_SIM = 0
    WIND_STABILITY = "'neutral'"
    AMBIENT_T = 27
    WIND_ROUGHNESS = 0.0002
    GEOMETRY_FILE = r"'E:\BR Team\P58\3DModel_P58.kfx'"
    VOL_BOX_MIN = "(-1e+030,-1e+030,-1e+030)"
    VOL_BOX_MAX = "(1e+030,1e+030,1e+030)"
    STOPPRO = 1
    DT_ROWCUM = 0.1
    DT_VOLSTOP = 10000
    LOCKEDPLANES = 2
    SPRAY_FILE = "()"
    JET_GAS_MASSFLOW = "()"
    JET_GAS_TIME = "()"
    GRID_PARAMS = "('Grow','N/A',0,1,0,0)"
    BLOCK_PARAMS = "(0,0,0,10,10,10,9999,9999,9999)"
    SUB_BLOCK = "(1,4000,1000,0,0,0,0,0,0)"
    POOL_PARAMS = "('Rectangle',0.01,0,0.005,373,288,0,1,0,0)"
    TIME_CONTROLS = "(1,1e+030,10,0,0,0,0,0)"
    HISTORY_POINTS = "'N/A'"
    SUBSEA_PARAMS = "('Gauss',1,1.3,0,0,0,0,'N/A','N/A')"
    RBM = "('Segment_1',70,790000,600,0,0,1,1e-005,0.00235,0,0,0,0,0,0,0)"
    XMLSPRAY = "(0.0005,0.001,0.0005,0.5,0,0,0,0)"
    EXPANDED = "(0,323.072,-5.073,0,0.103779,1,0,1,0,0,0)"
    VISTEMP = "('DEFAULT',0,0,0,0,0,0)"

    # Cabeçalho do arquivo SCN
    header = (
        "#Case_name                                                       data_type tmax jet_position            "
        "mstep jet_direction res_temperature jet_pressure jet_flowrate jet_diameter transient_jet "
        "jet_gas_components jet_gas_composition liquid_composition gas_cloud_sim wind_angle wind_10 "
        "wind_stability wind_Z0 ambient_T wind_roughness geometry_file geometry_min geometry_max "
        "vol_box_min vol_box_max stoppro dt_rowcum dt_volstop gridpoints lockedplanes spray_file "
        "jet_gas_massflow jet_gas_time grid_parameters block_parameters sub_block pool_parameters "
        "time_controls history_points subsea_parameters rbm xmlspray expanded_parameters vistemp_parameters\n"
    )

    print("\nIniciando geração de arquivos .scn...")
    for index, row in df.iterrows():
        # Obter nome do cenário da coluna apropriada
        if 'Cenário' in row:
            case_id = str(row['Cenário']).strip()
        elif 'cenário' in row:
            case_id = str(row['cenário']).strip()
        elif 'Cenario' in row:
            case_id = str(row['Cenario']).strip()
        else:
            for col in row.index:
                if 'V5412001' in str(row[col]) or 'V5135001' in str(row[col]):
                    case_id = str(row[col]).strip()
                    break
            else:
                case_id = f"Cenario_{index+1}"

        print(f"Processando linha {index+1}: {case_id}")

        path_line = f"'./{case_id}/{case_id}'"

        # Obter valores específicos da linha
        jet_position = format_xyz(row['jet_position']) if 'jet_position' in row else "(0,0,0)"
        jet_direction = format_xyz(row['jet_direction']) if 'jet_direction' in row else "(0,0,0)"
        if 'jet_flowrate' in row:
            try:
                jet_flowrate = float(str(row['jet_flowrate']).replace(',', '.'))
            except:
                jet_flowrate = 0.0
        else:
            jet_flowrate = 0.0
        res_temperature = row['res_temperature'] if 'res_temperature' in row and not pd.isna(row['res_temperature']) else RES_TEMP
        jet_pressure = row['jet_pressure'] if 'jet_pressure' in row and not pd.isna(row['jet_pressure']) else JET_PRESSURE
        diameter = None
        jet_diameter = 0.0
        if 'diameter' in row and not pd.isna(row['diameter']):
            try:
                diameter = float(str(row['diameter']).replace(',', '.').strip())
                jet_diameter = diameter
            except Exception:
                diameter = None
                jet_diameter = 0.0
        geometry_min = format_xyz(row['geometry_min']) if 'geometry_min' in row and not pd.isna(row['geometry_min']) else "(0,0,0)"
        geometry_max = format_xyz(row['geometry_max']) if 'geometry_max' in row and not pd.isna(row['geometry_max']) else "(0,0,0)"
        if 'transient_jet' in row:
            transient_jet = format_path(row['transient_jet'])
        else:
            transient_jet = "''"
        wind_angle = row['wind_angle'] if 'wind_angle' in row else 0
        wind_10 = row['wind_10'] if 'wind_10' in row else 0
        wind_stability = row['wind_stability'] if 'wind_stability' in row else "'neutral'"
        wind_Z0 = row['wind_Z0'] if 'wind_Z0' in row else 27
        ambient_T = row['ambient_T'] if 'ambient_T' in row else 27
        wind_roughness = row['wind_roughness'] if 'wind_roughness' in row else 0.0002
        lockedplanes = row['lockedplanes'] if 'lockedplanes' in row else 2
        spray_file = row['spray_file'] if 'spray_file' in row else "()"
        jet_gas_massflow = row['jet_gas_massflow'] if 'jet_gas_massflow' in row else "()"
        jet_gas_time = row['jet_gas_time'] if 'jet_gas_time' in row else "()"
        grid_parameters = row['grid_parameters'] if 'grid_parameters' in row else "('Grow','N/A',0,1,0,0)"
        gridpoints = row['gridpoints'] if 'gridpoints' in row else ""
        block_parameters = row['block_parameters'] if 'block_parameters' in row else "(0,0,0,10,10,10,9999,9999,9999)"
        sub_block = row['sub_block'] if 'sub_block' in row else "(1,4000,1000,0,0,0,0,0,0)"
        pool_parameters = row['pool_parameters'] if 'pool_parameters' in row else "('Rectangle',0.01,0,0.005,373,288,0,1,0,0)"
        time_controls = row['time_controls'] if 'time_controls' in row else "(1,1e+030,10,0,0,0,0,0)"
        history_points = row['history_points'] if 'history_points' in row else "'N/A'"
        subsea_parameters = row['subsea_parameters'] if 'subsea_parameters' in row else "('Gauss',1,1.3,0,0,0,0,'N/A','N/A')"
        rbm = row['rbm'] if 'rbm' in row else "('Segment_1',70,790000,600,0,0,1,1e-005,0.00235,0,0,0,0,0,0,0)"
        xmlspray = row['xmlspray'] if 'xmlspray' in row else "(0.0005,0.001,0.0005,0.5,0,0,0,0)"
        expanded_parameters = row['expanded_parameters'] if 'expanded_parameters' in row else "(0,323.072,-5.073,0,0.103779,1,0,1,0,0,0)"
        vistemp_parameters = row['vistemp_parameters'] if 'vistemp_parameters' in row else "('DEFAULT',0,0,0,0,0,0)"

        # Get module and set composition
        module = str(row['Module']).strip() if 'Module' in row else None
        jet_gas_components = "()"  # fallback
        jet_gas_composition = "()"
        if module and module in comp_sheets:
            comp_df = comp_sheets[module]
            cas_list = comp_df['CAS'].dropna().astype(str).tolist()
            mol_list = comp_df['Molar Amount'].dropna().astype(float).tolist()
            jet_gas_components = "(" + ",".join([f"'C:{cas}'" for cas in cas_list]) + ")"
            jet_gas_composition = "(" + ",".join([str(mol) for mol in mol_list]) + ")"

        # Checagem obrigatória: se diameter não foi definido, pode lançar erro ou pular cenário
        if diameter is None:
            print(f"❌ Erro: diameter não definido para o cenário {case_id}. Cenário ignorado.")
            continue

        # Construir linha de dados
        line = (
            f"{path_line:<65}"  # Case_name
            f"{'Jet_release_data:':<20}"  # data_type
            f"{TMAX:<25}"  # tmax
            f"{jet_position:<22}"  # jet_position
            f"{str(MSTEP):<8}"  # mstep
            f"{jet_direction:<18}"  # jet_direction
            f"{str(res_temperature):<16}"  # res_temperature
            f"{str(jet_pressure):<12}"  # jet_pressure
            f"{str(jet_flowrate):<12}"  # jet_flowrate
            f"{str(jet_diameter):<12}"  # jet_diameter
            f"{transient_jet:<60}"  # transient_jet
            f"{jet_gas_components:<80}"  # jet_gas_components
            f"{jet_gas_composition:<80}"  # jet_gas_composition
            f"{LIQUID_COMPOSITION:<40}"  # liquid_composition
            f"{str(GAS_CLOUD_SIM):<14}"  # gas_cloud_sim
            f"{str(wind_angle):<10}"  # wind_angle
            f"{str(wind_10):<10}"  # wind_10
            f"{wind_stability:<14}"  # wind_stability
            f"{str(wind_Z0):<10}"  # wind_Z0
            f"{str(ambient_T):<10}"  # ambient_T
            f"{str(wind_roughness):<14}"  # wind_roughness
            f"{GEOMETRY_FILE:<40}"  # geometry_file
            f"{geometry_min:<18}"  # geometry_min
            f"{geometry_max:<18}"  # geometry_max
            f"{VOL_BOX_MIN:<22}"  # vol_box_min
            f"{VOL_BOX_MAX:<22}"  # vol_box_max
            f"{str(STOPPRO):<10}"  # stoppro
            f"{str(DT_ROWCUM):<10}"  # dt_rowcum
            f"{str(DT_VOLSTOP):<10}"  # dt_volstop
            f"{str(gridpoints):<12}"  # gridpoints (now from Excel)
            f"{lockedplanes:<12}"  # lockedplanes
            f"{spray_file:<18}"  # spray_file
            f"{jet_gas_massflow:<18}"  # jet_gas_massflow
            f"{jet_gas_time:<18}"  # jet_gas_time
            f"{grid_parameters:<30}"  # grid_parameters
            f"{block_parameters:<30}"  # block_parameters
            f"{sub_block:<30}"  # sub_block
            f"{pool_parameters:<40}"  # pool_parameters
            f"{time_controls:<30}"  # time_controls
            f"{history_points:<12}"  # history_points
            f"{subsea_parameters:<40}"  # subsea_parameters
            f"{rbm:<60}"  # rbm
            f"{xmlspray:<30}"  # xmlspray
            f"{expanded_parameters:<40}"  # expanded_parameters
            f"{vistemp_parameters:<30}\n"  # vistemp_parameters
        )

        # Salvar arquivo SCN
        file_name = f"{case_id}.scn"
        file_path = os.path.join(output_dir, file_name)
        with open(file_path, 'w') as file:
            file.write(header)
            file.write(line)

        print(f"Arquivo gerado: {file_path}")

except Exception as e:
    print("Erro ao processar o arquivo:", e)
    sys.exit(1)

print("\nGeração de arquivos .scn concluída.")