import pandas as pd
import sys
import json
import re
import os
from pathlib import Path
from datetime import datetime

def merge_arch_planner(target_folder_name, manual_hw_name=None):
    # 1. Setup Paths
    base_path = Path(__file__).parent.absolute()
    input_dir = base_path / target_folder_name
    source_dir = input_dir / "ArchPlannerOutput"
    # Adjusted path to match your environment
    hw_config_dir = Path(r"C:\Users\andychanyan\ceva-Studio\workspace_np_v25.0.0.GA\.ceva-workspace-vscode\hw-config")
    
    command_log_path = source_dir / "command_log.txt"
    config_file = source_dir / "config" / "selected_values.json"
    summary_table_csv = source_dir / "summary_table.csv"
    
    if not source_dir.exists():
        print(f"Error: Could not find directory: {source_dir}")
        return

    # --- MODEL NAME TRANSLATION DICTIONARY ---
    translation_dict = {
        "ad01_int8_qo.onnx" : "TinyML Perf Anomaly Detection - ONNX Variant", 
        "kws_ref_model_qo.onnx" : "TinyML Perf Key Word Spotting - ONNX Variant",
        "pretrainedResnet_quant_qo.onnx" :"TinyML Perf ResNet18 - ONNX Variant", 
        "vww_96_int8_qo.onnx" : "TinyML Perf VWW - Visual Wake Word - ONNX Variant", 
        "ad01_int8.tflite" : "TinyML Perf Anomaly Detection", 
        "efficientnet-lite0-int8.tflite" : "EfficientNet lite0", 
        "keras_mobilenet_v3_small_075_quantized_model.tflite" : "Mobilenet V3 small 075", 
        "kws_micronet_l.tflite" : "KWS Micronet - Key Word Spotting",
        "kws_ref_model.tflite" : "TinyML Perf Key Word Spotting", 
        "pretrainedResnet_quant.tflite" : "TinyML Perf ResNet18", 
        "squeezenet1.1-7_full_integer_quant.tflite" : "Squeezenet1.1",
        "vww_96_int8.tflite" : "TinyML Perf VWW - Visual Wake Word"
    }

    # 2. DATA COLLECTION
    total_cycles, sum_op_cycles, p_12nm, p_22nm, p_40nm = 1, 0, 0, 0, 0
    clock_hz = 400000000  # 400 MHz
    config_data, current_processor, npu_params_str = {}, "N/A", "N/A"
    data_type = "N/A"

    # A. Cycles & Power
    if summary_table_csv.exists():
        try:
            df_sum = pd.read_csv(summary_table_csv)
            if 'Cycle Count' in df_sum.columns:
                total_cycles = int(df_sum['Cycle Count'].iloc[0])
            
            col_power = 'Average Dynamic Power (mW/MHz) 400MHz'
            if col_power in df_sum.columns:
                p_12nm = round(df_sum.iloc[0][col_power] * 400, 2)
                p_22nm = round(df_sum.iloc[1][col_power] * 400, 2) if len(df_sum) > 1 else 0
                p_40nm = round(df_sum.iloc[2][col_power] * 400, 2) if len(df_sum) > 2 else 0
        except: pass

    # B. Op Cycles
    full_graph_file = next(source_dir.glob("*FullGraph.csv"), None)
    if full_graph_file:
        try:
            df_graph = pd.read_csv(full_graph_file)
            if 'Cycles' in df_graph.columns: sum_op_cycles = df_graph['Cycles'].sum()
        except: pass

    # C. Processor Config
    if config_file.exists():
        try:
            with open(config_file, 'r', encoding='utf-8') as f:
                config_data = json.load(f)
            current_processor = manual_hw_name if manual_hw_name else config_data.get("processor", "N/A")
            dtcm = config_data.get("DTCM_Arena_bytes", "N/A")
            bw_cycles = config_data.get("bandwidth_cycles", "N/A")
            data_type = config_data.get("data_type", "N/A")
            npu_params_str = f"Proc: {current_processor}, DTCM: {dtcm}, BW: {bw_cycles}"
        except: pass

    # D. Log Parsing
    sim_choice, network_display_name = "Ceva-NeuPro-M", "Unknown Network"
    log_raw_text = "Log file not found."
    if command_log_path.exists():
        try:
            log_raw_text = command_log_path.read_text()
            if "NeuPro-Nano" in log_raw_text: sim_choice = "NeuPro-Nano"
            match = re.search(r"([^\s\\]+\.(?:onnx|tflite))", log_raw_text)
            if match: network_display_name = translation_dict.get(match.group(1), match.group(1))
        except: pass

    # 3. CALCULATIONS
    ips_value = int(round(clock_hz / total_cycles, 0))
    overhead_cycles = max(0, total_cycles - sum_op_cycles)
    lat_pct = round((overhead_cycles / total_cycles) * 100, 2)
    
    if data_type != "N/A":
        bit_match = re.search(r'\d+', str(data_type))
        bit_val = bit_match.group(0) if bit_match else str(data_type)
        sdk_support = f"Yes ({bit_val} bits)"
    else:
        sdk_support = "No"

    # --- FILENAME ---
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    clean_net_name = str(network_display_name).replace(' ', '_').replace('.', '_')
    output_filename = f"{sim_choice}_{clean_net_name}_{timestamp}.xlsx"
    output_path = input_dir / output_filename

    # 4. EXCEL GENERATION
    try:
        writer = pd.ExcelWriter(output_path, engine='xlsxwriter')
        workbook = writer.book

        dash = workbook.add_worksheet('Dashboard')
        dash.activate()
        dash.set_column('A:A', 40) 
        dash.set_column('B:B', 70) 

        title_fmt = workbook.add_format({'bold': True, 'font_size': 14, 'font_color': '#2E75B6'})
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC', 'border': 1})
        val_fmt = workbook.add_format({'border': 1, 'text_wrap': True, 'align': 'left'})

        dash.write('A1', 'CEVA ARCHPLANNER ANALYSIS SUMMARY', title_fmt)
        
        summary_items = [
            ["Date/Time", datetime.now().strftime("%m/%d/%Y %H:%M")],
            ["Simulator Choice", sim_choice],
            ["Hardware Configuration", current_processor],
            ["Interpreted Network Name", network_display_name],
            ["NPU Parameters (Detailed)", npu_params_str],
            ["IPS (Inferences Per Second)", ips_value],
            ["Scheduler Latency (%)", f"{lat_pct}%"],
            ["SDK Support for Quantization", sdk_support],
            ["Power Usage 12nm (mW)", p_12nm],
            ["Power Usage 22nm (mW)", p_22nm],
            ["Power Usage 40nm (mW)", p_40nm]
        ]
        
        for r, (k, v) in enumerate(summary_items, start=2):
            dash.write(r, 0, k, header_fmt)
            dash.write(r, 1, v, val_fmt)

        # --- POWER BAR CHART ---
        bar = workbook.add_chart({'type': 'column'})
        bar.add_series({
            'name': 'Power (mW)',
            'categories': ['Dashboard', 10, 0, 12, 0], 
            'values':     ['Dashboard', 10, 1, 12, 1],
            'fill': {'color': '#4F81BD'},
            'data_labels': {'value': True}
        })
        bar.set_title({'name': 'Estimated Dynamic Power Consumption'})
        dash.insert_chart('D1', bar)

        # --- SCHEDULER LATENCY PIE CHART ---
        # Data source for the pie chart (Hidden in Columns H and I)
        dash.write('H1', 'Cycle Category')
        dash.write('H2', 'Compute Cycles')
        dash.write('H3', 'Scheduler Latency')
        dash.write('I1', 'Count')
        dash.write('I2', sum_op_cycles)
        dash.write('I3', overhead_cycles)

        pie = workbook.add_chart({'type': 'pie'})
        pie.add_series({
            'name':       'Scheduler Latency Analysis',
            'categories': ['Dashboard', 1, 7, 2, 7], # H2:H3
            'values':     ['Dashboard', 1, 8, 2, 8], # I2:I3
            'points': [
                {'fill': {'color': '#4F81BD'}}, # Compute (Blue)
                {'fill': {'color': '#C0504D'}}, # Latency (Red)
            ],
            'data_labels': {
                'percentage': True, 
                'category': True, 
                'position': 'outside_end',
                'leader_lines': True
            }
        })
        pie.set_title({'name': f'Latency Impact: {lat_pct}%'})
        dash.insert_chart('D16', pie) # Placed below the Bar Chart

        # Tabs: Command Log, Config, and HW Specs
        log_sheet = workbook.add_worksheet('Command_Log')
        log_fmt = workbook.add_format({'font_name': 'Courier New', 'font_size': 9})
        for i, line in enumerate(log_raw_text.splitlines()):
            log_sheet.write(i, 0, line, log_fmt)

        pd.json_normalize(config_data).to_excel(writer, sheet_name="CONFIG_SelectedValues", index=False)
        
        hw_data = []
        tdf_file = hw_config_dir / f"{current_processor}.tdf"
        if not tdf_file.exists(): 
            tdf_file = hw_config_dir / f"{str(current_processor).split('_')[0]}.tdf"
        
        if tdf_file.exists():
            content = tdf_file.read_text(encoding='utf-8', errors='ignore')
            matches = re.findall(r'name="([^"]+)"\s+value="([^"]*)"', content)
            hw_data = [{"Parameter": m[0], "Value": m[1]} for m in matches]
        pd.DataFrame(hw_data).to_excel(writer, sheet_name="Hardware_Specs", index=False)

        # Dynamic CSV import for all other result files (INCLUDING summary_table)
        for f in source_dir.iterdir():
            if f.is_file() and f.suffix.lower() == '.csv':
                try:
                    # Clean the sheet name (max 31 chars)
                    sheet_name = f.stem[:31]
                    pd.read_csv(f).to_excel(writer, sheet_name=sheet_name, index=False)
                except: continue

        writer.close()
        print(f"SUCCESS! Created Analysis: {output_filename}")

    except Exception as e:
        print(f"Critical Error: {e}")

if __name__ == "__main__":
    if len(sys.argv) > 1:
        merge_arch_planner(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None)
