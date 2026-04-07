import pandas as pd
from decimal import Decimal, ROUND_HALF_UP
import re
import streamlit as st
import io



def find_exact_sum(target, available_values):
    # Safely convert the target to an integer
    target = int(float(target))
    
    # Safely clean and convert available values
    clean_values = set()
    for v in available_values:
        try:
            # Convert to float first (to handle strings like "5.0"), then to int
            num = int(float(v))
            if num > 0:
                clean_values.add(num)
        except (ValueError, TypeError):
            # If the value is text, NaN, or completely invalid, just skip it
            continue
            
    # Sort descending: trying the biggest chunks first is much faster
    available_values = sorted(clean_values, reverse=True)
    
    best_solution = None
    min_coins = float('inf')

    def backtrack(remaining, current_index, current_combo, current_coin_count):
        nonlocal best_solution, min_coins
        
        # If we hit exactly 0, we found a valid combination
        if remaining == 0:
            if current_coin_count < min_coins:
                min_coins = current_coin_count
                best_solution = current_combo.copy()
            return
        
        # If we run out of values to try, stop
        if current_index >= len(available_values):
            return
        
        val = available_values[current_index]
        
        # MATHEMATICAL PRUNING
        if current_coin_count + (remaining // val) >= min_coins:
            return

        # Try using the maximum possible amount of the current value, down to 0
        max_amount = remaining // val
        for count in range(max_amount, -1, -1):
            if count > 0:
                current_combo[val] = count
            else:
                if val in current_combo:
                    del current_combo[val]
            
            backtrack(remaining - (count * val), current_index + 1, current_combo, current_coin_count + count)
            
            if val in current_combo:
                del current_combo[val]

    # Start the calculation
    backtrack(target, 0, {}, 0)
    return best_solution


def normalize_val(val):
    if pd.isna(val) or val == "":
        return ""
    
    s_val = str(val).strip()
    
    try:
        d = Decimal(s_val)
    except:
        match = re.match(r"([-+]?\d*\.?\d+)", s_val)
        if match:
            d = Decimal(match.group(1))
        else:
            return s_val 
    
    rounded = d.quantize(Decimal('1'), rounding=ROUND_HALF_UP)
    return str(int(rounded))

def process_file(file):
    df_facturar = pd.read_excel(file, dtype=str, skiprows=3)
    #df_facturar.columns = ['Raw_Value']

    print(df_facturar.head())
    print(df_facturar.columns)
    df_mapping = pd.read_csv('productos_processed.csv', dtype=str)

    df_facturar['Normalized_Value'] = df_facturar['Delivery Charge']#.apply(normalize_val)
    counts = df_facturar['Normalized_Value'].value_counts().reset_index()
    counts.columns = ['Valor', 'Recuento']


    #df_mapping['Valor_List'] = df_mapping['Valor'].astype(str).str.split(',')
    #df_exploded = df_mapping.explode('Valor_List')
    #df_exploded['Normalized_Value'] = df_exploded['Valor_List'].apply(normalize_val)

    mapping_lookup = df_mapping.groupby('PRECIO_VENTA')['CODPROD'].apply(
        lambda x: ', '.join(sorted(set([str(c).strip() for c in x if pd.notna(c)])))
    ).reset_index()
    mapping_lookup.columns = ['Valor', 'Codigo']
    mapping_lookup['Codigo'] = mapping_lookup['Codigo'].str.replace('=', '', regex=False).str.replace('"', '', regex=False)


    result = pd.merge(counts, mapping_lookup, on='Valor', how='left')
    sum_empty = result.loc[result['Codigo'].isna(), 'Valor'].sum()
    result['Codigo'] = result['Codigo'].fillna('-')
    list_of_values = result.loc[result['Codigo'] != '-', 'Valor']
    result = result[['Valor', 'Codigo', 'Recuento']]
    
    solution = find_exact_sum(sum_empty, list_of_values)

    #############
    valor_numeric = pd.to_numeric(result['Valor'], errors='coerce').fillna(0)
    recuento_numeric = pd.to_numeric(result['Recuento'], errors='coerce').fillna(0)

    # Sum of (Valor * Recuento)
    total_sum = (valor_numeric * recuento_numeric).sum()

    # Create the final row
    new_row = pd.DataFrame({
        'Valor': [f"TOTAL: {total_sum}"], 
        'Codigo': [''],       # Empty string for other columns
        'Recuento': ['']
    })
    
    sum_new_row = pd.DataFrame({
        'Valor': [f"TOTAL Sin Codigo:"], 
        'Codigo': [f"{sum_empty}"],       # Empty string for other columns
        'Recuento': ['']
    })

    #Append the row to the original DataFrame
    # 'ignore_index=True' ensures the row numbers continue perfectly from the old data
    result = pd.concat([result, new_row], ignore_index=True)
    result = pd.concat([result, sum_new_row], ignore_index=True)
    
    new_rows = pd.DataFrame(list(solution.items()), columns=['Codigo', 'Recuento'])

    result = pd.concat([result, new_rows], ignore_index=True)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        result.to_excel(writer, index=False, sheet_name='Sheet1')
        
        workbook  = writer.book
        worksheet = writer.sheets['Sheet1']
        
        # Crea un formato testo (Text format '@')
        text_format = workbook.add_format({'num_format': '@'})
        
        # Applica il formato alla colonna B (che è 'Codigo' se Valor è A e Recuento è C)
        # Excel usa indici base-0 per le colonne (A=0, B=1, C=2)
        # Impostiamo anche una larghezza colonna per leggibilità
        worksheet.set_column('A:A', 20)              # Colonna Valor
        worksheet.set_column('B:B', 20, text_format) # Colonna Codigo (FORZATA A TESTO)
        worksheet.set_column('C:C', 15)              # Colonna Recuento

    output.seek(0)
    return result, output.getvalue(), total_sum




# --- STREAMLIT UI SETUP ---
st.set_page_config(page_title="Report Transformer", layout="centered")

st.title("📊 Output Mappings")
st.markdown("""
Upload your **TransactionReport Converted** file below. 
This tool will generate a report mapping values to product codes.
""")

# File Uploader (Drag and Drop)
uploaded_file = st.file_uploader("Drag and drop CSV file here", type=["xlsx", "csv"])

if uploaded_file is not None:
    st.success("File uploaded successfully!")
    
    # Process
    df_result, xlsx_data, total_val = process_file(uploaded_file)
    
    if df_result is not None:
        # Show Preview
        st.subheader("Preview of Generated Output Mapping")
        st.metric(label="Calculated Total", value=f"{total_val:,.2f}")
        st.dataframe(df_result.head(10))
        
        # Generate CSV String
        #csv_data = generate_excel_bytes(df_result)
        
        # Download Button
        st.download_button(
            label="📥 Download Output Mappings",
            data=xlsx_data,
            file_name="Output_Mappings.xlsx",
            mime="text/csv"
        )