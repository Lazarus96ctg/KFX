import pandas as pd
import os
import re
import sys

# Caminho para o Excel
excel_path = r'C:\Users\LENOVO\Documents\ABS\KFX\Scripts\Controle de simulação_teste_automatização.xlsx'

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
        nrows=8    # Apenas 8 linhas de dados
    )
    
    print(f"Planilha carregada com {len(df)} linhas e {len(df.columns)} colunas")
    print("Colunas encontradas:", df.columns.tolist())
    
    # Valores fixos para colunas não presentes na planilha
    TMAX = "1e+030"
    MSTEP = 900000
    RES_TEMP = 72
    JET_PRESSURE = 39.42
    JET_DIAMETER = 2
    JET_GAS_COMPONENTS = "('C:7727379','C:124389','C:74828','C:74840','C:74986','C:106978','C:109660','C:110543','C:142825','C:111659','C:111842','C:124185')"
    JET_GAS_COMPOSITION = "(0.977456,0.0508478,59.1947,12.5749,10.3382,7.87036,4.20464,2.1307,1.4402,0.791865,0.308222,0.11836)"
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
    LOCKEDPLANES = "'N/A'"
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

    # Geração dos arquivos SCN
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
            # Tenta encontrar a coluna de cenário pelo conteúdo
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
        
        # Obter vazão do jato
        if 'jet_flowrate' in row:
            try:
                jet_flowrate = float(str(row['jet_flowrate']).replace(',', '.'))
            except:
                jet_flowrate = 0.0
        else:
            jet_flowrate = 0.0
        
        # Obter caminho do arquivo transiente
        if 'transient_jet' in row:
            transient_jet = format_path(row['transient_jet'])
        else:
            transient_jet = "''"
        
        # Valores padrão para outras colunas
        wind_angle = row['wind_angle'] if 'wind_angle' in row else 0
        wind_10 = row['wind_10'] if 'wind_10' in row else 0.0
        geometry_min = format_xyz(row['geometry_min']) if 'geometry_min' in row else "(0,0,0)"
        geometry_max = format_xyz(row['geometry_max']) if 'geometry_max' in row else "(0,0,0)"
        gridpoints = row['gridpoints'] if 'gridpoints' in row else 150000
        
        # Construir linha de dados
        line = (
            f"{path_line:<65} 'Jet_release_data:' {TMAX} {jet_position} "
            f"{MSTEP} {jet_direction} {RES_TEMP} {JET_PRESSURE} {jet_flowrate} "
            f"{JET_DIAMETER} {transient_jet} {JET_GAS_COMPONENTS} {JET_GAS_COMPOSITION} "
            f"{LIQUID_COMPOSITION} {GAS_CLOUD_SIM} {wind_angle} {wind_10} "
            f"{WIND_STABILITY} {AMBIENT_T} {wind_10} {WIND_ROUGHNESS} "
            f"{GEOMETRY_FILE} {geometry_min} {geometry_max} "
            f"{VOL_BOX_MIN} {VOL_BOX_MAX} {STOPPRO} {DT_ROWCUM} {DT_VOLSTOP} "
            f"{gridpoints} 2 {LOCKEDPLANES} {SPRAY_FILE} {JET_GAS_MASSFLOW} "
            f"{JET_GAS_TIME} {GRID_PARAMS} {BLOCK_PARAMS} {SUB_BLOCK} "
            f"{POOL_PARAMS} {TIME_CONTROLS} {HISTORY_POINTS} {SUBSEA_PARAMS} "
            f"{RBM} {XMLSPRAY} {EXPANDED} {VISTEMP}\n"
        )

        output_path = os.path.join(output_dir, f"{case_id}.scn")
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(header)
            f.write(line)
        print(f"✓ Arquivo gerado: {output_path}")

    print(f"\n✅ {len(df)} arquivos .scn gerados com sucesso em: {output_dir}")

except Exception as e:
    print(f"❌ Erro crítico: {str(e)}")
    import traceback
    traceback.print_exc()
    sys.exit(1)