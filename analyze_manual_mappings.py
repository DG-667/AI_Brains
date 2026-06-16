import pandas as pd

file_path = r'C:\Users\DenisG\Documents\Workplace\Data\GFK\GFK data TTL.xlsx'
df = pd.read_excel(file_path, sheet_name='Data')
df.columns = df.columns.str.strip()

ALLOWED_SEGMENTS = {
    'Apple', 'Basic Celeron', 'Basic i3R3', 'Basic i5R5', 'Basic i7R7', 'Basic i9R9', 
    'Basic Pentium Athlon', 'Gaming i5R5 50series', 'Gaming i5R5 60series', 'Gaming i7R7 50series', 
    'Gaming i7R7 60series', 'Gaming i7R7 70series', 'Gaming i9R9 60series', 'Gaming i9R9 70series', 
    'Gaming i9R9 80+series', 'Gaming Others', 'Premium i3R3', 'Premium i5R5', 'Premium i7R7', 
    'Premium i9R9', 'Snapdragon'
}

ALLOWED_SUBSEGMENTS = {
    'Apple', 'Basic Celeron', 'Basic Pentium Athlon', 'Basic R3', 'Basic i3', 'Basic R5', 'Basic i5', 
    'Basic R7', 'Basic i7', 'Basic R9', 'Basic i9', 'Premium i3R3', 'Premium i5R5', 'Premium i7R7', 
    'Premium i9R9', 'SnapDr Laptops', 'Gaming i5R5 3050', 'Gaming i7R7 3050', 'Gaming i5R5 3060', 
    'Gaming i7R7 3060', 'Gaming i7R7 3070+', 'Gaming i5R5 4050', 'Gaming i7R7 4050', 'Gaming i5R5 4060', 
    'Gaming i7R7 4060', 'Gaming i7R7 4070+', 'Gaming i9R9 4070+', 'Gaming i9R9 4080+', 'Gaming i5R5 5050', 
    'Gaming i7R7 5050', 'Gaming i5R5 5060', 'Gaming i7R7 5060', 'Gaming i9R9 5060', 'Gaming i7R7 5070+', 
    'Gaming i9R9 5070+', 'Gaming Others'
}

def get_mapping(brand, processor, gaming_pc, model_name, gpu_model):
    brand, processor, gaming_pc = str(brand).lower(), str(processor).lower(), str(gaming_pc).lower()
    model_name, gpu_model = str(model_name).lower(), str(gpu_model).lower()
    if brand == 'apple': return 'Apple', 'Apple'
    if 'snapdragon' in processor: return 'Snapdragon', 'SnapDr Laptops'
    cpu_seg, cpu_sub = '', ''
    if 'core i3' in processor or 'core 3' in processor or 'core ultra 3' in processor: cpu_seg, cpu_sub = 'i3R3', 'i3'
    elif 'ryzen 3' in processor or 'ryzen ai 3' in processor or 'ryzen ai 300' in processor: cpu_seg, cpu_sub = 'i3R3', 'R3'
    elif 'core i5' in processor or 'core 5' in processor or 'core ultra 5' in processor: cpu_seg, cpu_sub = 'i5R5', 'i5'
    elif 'ryzen 5' in processor or 'ryzen ai 5' in processor: cpu_seg, cpu_sub = 'i5R5', 'R5'
    elif 'core i7' in processor or 'core 7' in processor or 'core ultra 7' in processor: cpu_seg, cpu_sub = 'i7R7', 'i7'
    elif 'ryzen 7' in processor or 'ryzen ai 7' in processor: cpu_seg, cpu_sub = 'i7R7', 'R7'
    elif 'core i9' in processor or 'core 9' in processor or 'core ultra 9' in processor: cpu_seg, cpu_sub = 'i9R9', 'i9'
    elif 'ryzen 9' in processor or 'ryzen ai 9' in processor: cpu_seg, cpu_sub = 'i9R9', 'R9'
    elif 'celeron' in processor: cpu_seg, cpu_sub = 'Celeron', 'Celeron'
    elif 'pentium' in processor or 'athlon' in processor: cpu_seg, cpu_sub = 'Pentium Athlon', 'Pentium Athlon'
    gaming_models = ['legion', 'rog', 'tuf', 'omen', 'nitro', 'alienware', 'predator', 'loq']
    is_gaming = (gaming_pc == 'gaming') or any(m in model_name for m in gaming_models)
    if is_gaming:
        if not cpu_seg: return 'Gaming Others', 'Gaming Others'
        gpu_seg, gpu_sub = '', ''
        if '3050' in gpu_model: gpu_seg, gpu_sub = '50series', '3050'
        elif '3060' in gpu_model: gpu_seg, gpu_sub = '60series', '3060'
        elif any(x in gpu_model for x in ['3070', '3080', '3090']): gpu_seg, gpu_sub = '70series', '3070+'
        elif '4050' in gpu_model: gpu_seg, gpu_sub = '50series', '4050'
        elif '4060' in gpu_model: gpu_seg, gpu_sub = '60series', '4060'
        elif '4070' in gpu_model: gpu_seg, gpu_sub = '70series', '4070+'
        elif any(x in gpu_model for x in ['4080', '4090']): gpu_seg, gpu_sub = '80+series', '4080+'
        elif '5050' in gpu_model: gpu_seg, gpu_sub = '50series', '5050'
        elif '5060' in gpu_model: gpu_seg, gpu_sub = '60series', '5060'
        elif any(x in gpu_model for x in ['5070', '5080', '5090']): gpu_seg, gpu_sub = '70series', '5070+'
        if not gpu_seg: return 'Gaming Others', 'Gaming Others'
        return f'Gaming {cpu_seg} {gpu_seg}', f'Gaming {cpu_seg} {gpu_sub}'
    premium_models = ['zenbook', 'spectre', 'xps', 'yoga', 'envy', 'elitebook', 'swift']
    if any(m in model_name for m in premium_models):
        if not cpu_seg: return '', ''
        return f'Premium {cpu_seg}', f'Premium {cpu_seg}'
    if not cpu_seg: return '', ''
    return f'Basic {cpu_seg}', f'Basic {cpu_sub}'

results = []
for idx, row in df.iterrows():
    seg, sub = get_mapping(row.get('Brand',''), row.get('Processor',''), row.get('Gaming PCs', ''), row.get('Model Name',''), row.get('GPU Model', ''))
    if seg not in ALLOWED_SEGMENTS: seg = ''
    if sub not in ALLOWED_SUBSEGMENTS: sub = ''
    actual_seg = str(row.get('Tech_Segment', '')).strip()
    actual_sub = str(row.get('Tech_SubSegment', '')).strip()
    if actual_seg == 'nan': actual_seg = ''
    if actual_sub == 'nan': actual_sub = ''
    if seg == '' and actual_seg != '':
        results.append({
            'Processor': str(row.get('Processor', '')),
            'GPU': str(row.get('GPU Model', '')),
            'Model': str(row.get('Model Name', '')),
            'Brand': str(row.get('Brand', '')),
            'User_Seg': actual_seg,
            'User_Sub': actual_sub
        })

res_df = pd.DataFrame(results)
if res_df.empty:
    print('NO MANUAL CORRECTIONS FOUND.')
else:
    print(f'FOUND {len(res_df)} MANUAL CORRECTIONS.')
    summary = res_df.groupby(['User_Seg', 'User_Sub', 'Processor', 'GPU']).size().reset_index(name='count')
    print(summary.sort_values('count', ascending=False).to_string(index=False))


