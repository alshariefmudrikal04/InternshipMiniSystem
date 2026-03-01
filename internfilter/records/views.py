from django.shortcuts import render
import pandas as pd

file_path = 'data/202507 DATABASE.xlsx'
df = pd.read_excel(file_path)

def index(request):

    data = None
    selected_column = None

    if request.method == "POST":
        name = request.POST.get('name')
        selected_column = request.POST.get('column')

        filtered = df[df['FULLNAME'].str.contains(name, case=False, na=False)]

        if selected_column:
            filtered = filtered[['FULLNAME', selected_column]]

        data = filtered.to_dict('records')

    return render(request, 'records/index.html', {
        'data': data,
        'columns': df.columns.tolist(),
        'selected_column': selected_column
    })