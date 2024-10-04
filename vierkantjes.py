import os
from flask import Flask, render_template, request, redirect, url_for, send_file
import pandas as pd
import matplotlib
matplotlib.use('Agg')  # Use 'Agg' for non-GUI backend
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import openpyxl
import numpy as np

print(f"Using openpyxl version: {openpyxl.__version__}")

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
    df = pd.read_excel(file_path, engine='openpyxl')

    # Extract the data for distance groups and modal split
    distance_groups = df[['Afstandsklasse', 'Totaal']].dropna()
    modal_split = df[['Te voet', 'Fiets', 'OV', 'Auto (+overig)']]

    # Normalize the total proportions of distance groups
    total_proportions = distance_groups['Totaal'].sum()
    distance_groups['Totaal'] = distance_groups['Totaal'] / total_proportions

    # Create a square plot, smaller plot but add more space around it
    fig, ax = plt.subplots(figsize=(9, 9))

    y_offset = 0  # Initialize y offset for distance groups
    group_centers = []

    for idx, row in distance_groups.iterrows():
        distance_label = row['Afstandsklasse']
        total_proportion = row['Totaal']

        if pd.isna(distance_label):
            continue

        group_height = total_proportion - MODE_SPACING
        modal_split_proportions = modal_split.iloc[idx]
        total_split = modal_split_proportions.sum()
        modal_split_proportions = modal_split_proportions / total_split

        group_centers.append(y_offset + group_height / 2)

        x_offset = 0
        for mode, proportion in modal_split_proportions.items():
            width = proportion - MODE_SPACING
            ax.add_patch(plt.Rectangle((x_offset, y_offset), width, group_height, 
                                       color=colors[mode], edgecolor='none'))
            x_offset += proportion

        y_offset += total_proportion

    ax.set_yticks(group_centers)
    ax.set_yticklabels(distance_groups['Afstandsklasse'], fontsize=12, ha='right')

    ax.set_xticks([0.0, 0.2, 0.4, 0.6, 0.8, 1.0])
    ax.set_xticklabels([f'{int(i*100)}%' for i in [0.0, 0.2, 0.4, 0.6, 0.8, 1.0]], fontsize=12)

    # Set limits to provide more space for labels
    ax.set_ylim(0, 1)  # Adjust as needed based on total proportions

    # Adjust subplot parameters for better spacing
    plt.subplots_adjust(left=0.3, right=0.85, top=0.9, bottom=0.2)

    # Ensure square aspect ratio
    ax.set_aspect('equal', adjustable='box')

    # Add legend for transport modes
    handles = [plt.Rectangle((0, 0), 1, 1, color=color) for color in colors.values()]
    labels = list(colors.keys())
    ax.legend(handles, labels, loc='upper left', bbox_to_anchor=(1.05, 1), frameon=False)

    # Save the plot to the output path
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