import os
from flask import Flask, render_template, request, redirect, url_for, send_file
import pandas as pd
import matplotlib
matplotlib.use('Agg')  # Use 'Agg' for non-GUI backend
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker

import openpyxl
print(f"Using openpyxl version: {openpyxl.__version__}")


# Set matplotlib to non-interactive backend
matplotlib.use('Agg')

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads/'
app.config['PLOT_FOLDER'] = 'static/plots/'

# Ensure upload and plot directories exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['PLOT_FOLDER'], exist_ok=True)

# Define transport mode colors
colors = {
    'Te voet': '#70c282',  # green
    'Fiets': '#e62d73',    # pink
    'OV': '#ffd403',       # yellow
    'Auto (+overig)': '#038aa1'  # teal
}
MODE_SPACING = 0.005

def create_proportional_plot(file_path, output_path):
    df = pd.read_excel(file_path, engine='openpyxl')  # Specify openpyxl as the engine

    distance_groups = df[['Afstandsklasse', 'Totaal']].dropna()
    modal_split = df[['Te voet', 'Fiets', 'OV', 'Auto (+overig)']]

    total_proportions = distance_groups['Totaal'].sum()
    distance_groups['Totaal'] = distance_groups['Totaal'] / total_proportions

    fig, ax = plt.subplots(figsize=(8, 8))
    x_offset = 0
    spacing = 0.005

    group_centers = []

    for idx, row in distance_groups.iterrows():
        distance_label = row['Afstandsklasse']
        total_proportion = row['Totaal']

        if pd.isna(distance_label):
            continue

        group_width = total_proportion - spacing
        modal_split_proportions = modal_split.iloc[idx]
        total_split = modal_split_proportions.sum()
        modal_split_proportions = modal_split_proportions / total_split

        group_centers.append(x_offset + group_width / 2)
        y_offset = 0
        for mode, proportion in modal_split_proportions.items():
            height = proportion - MODE_SPACING
            ax.add_patch(plt.Rectangle((x_offset, y_offset), group_width, height, color=colors[mode], edgecolor='none'))
            y_offset += proportion

        x_offset += total_proportion

    ax.set_xlim([0, 1])
    ax.set_ylim([0, 1])
    ax.set_aspect('equal')
    ax.set_xticks(group_centers)
    ax.set_xticklabels(distance_groups['Afstandsklasse'], fontsize=12, ha='center')
    ax.set_yticks([0.0, 0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 1.0])
    ax.set_yticklabels([f'{int(i*100)}%' for i in [0.0, 0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 1.0]], fontsize=12)

    ax.xaxis.set_major_locator(ticker.MultipleLocator(0.1))
    ax.grid(which='major', axis='x', linestyle=':', linewidth=1, color='gray')
    ax.yaxis.set_major_locator(ticker.MultipleLocator(0.1))
    ax.grid(which='major', axis='y', linestyle=':', linewidth=1, color='gray')

    handles = [plt.Rectangle((0, 0), 1, 1, color=color) for color in colors.values()]
    labels = list(colors.keys())
    ax.legend(handles, labels, loc='upper left', bbox_to_anchor=(1.05, 1), frameon=False)

    plt.savefig(output_path, bbox_inches='tight')
    plt.close()

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return redirect(request.url)
    file = request.files['file']
    if file.filename == '' or file.filename.startswith('~$'):
        return redirect(request.url)
    if file:
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(file_path)
        output_plot = os.path.join(app.config['PLOT_FOLDER'], f'plot_{file.filename}.png')
        create_proportional_plot(file_path, output_plot)
        return send_file(output_plot, as_attachment=True)

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)