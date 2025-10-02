import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from scipy.interpolate import griddata
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

class ContourPlotGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Contour Plot Generator")
        self.root.geometry("516x600")
        
        # Data storage
        self.data = None
        self.filepath = None
        
        # Create GUI elements
        self.setup_gui()
        
    def setup_gui(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # File selection frame
        file_frame = ttk.LabelFrame(main_frame, text="File Selection", padding="5")
        file_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # File path display
        self.filepath_var = tk.StringVar(value="No file selected")
        ttk.Label(file_frame, text="Selected file:").grid(row=0, column=0, sticky=tk.W)
        ttk.Label(file_frame, textvariable=self.filepath_var, foreground="blue").grid(row=1, column=0, sticky=tk.W)
        
        # Browse file button
        ttk.Button(main_frame, text="Browse Excel File", command=self.browse_file).grid(row=1, column=0, columnspan=2, pady=10)
        
        # Sheet selection
        sheet_frame = ttk.LabelFrame(main_frame, text="Sheet Selection", padding="5")
        sheet_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Label(sheet_frame, text="Sheet name:").grid(row=0, column=0, sticky=tk.W)
        self.sheet_var = tk.StringVar(value="Sheet1")
        self.sheet_combo = ttk.Combobox(sheet_frame, textvariable=self.sheet_var, width=20)
        self.sheet_combo.grid(row=0, column=1, padx=(10, 0))
        
        # Reloads data when sheet_var is changed
        self.sheet_var.trace_add("write",self.load_data)
        
        # Column selection frame
        col_frame = ttk.LabelFrame(main_frame, text="Column Selection", padding="5")
        col_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Column dropdowns
        ttk.Label(col_frame, text="X Column:").grid(row=0, column=0, sticky=tk.W)
        self.x_col_var = tk.StringVar()
        self.x_col_combo = ttk.Combobox(col_frame, textvariable=self.x_col_var, width=15)
        self.x_col_combo.grid(row=0, column=1, padx=5)
        
        ttk.Label(col_frame, text="Y Column:").grid(row=0, column=2, sticky=tk.W)
        self.y_col_var = tk.StringVar()
        self.y_col_combo = ttk.Combobox(col_frame, textvariable=self.y_col_var, width=15)
        self.y_col_combo.grid(row=0, column=3, padx=5)
        
        ttk.Label(col_frame, text="Z Column:").grid(row=1, column=0, sticky=tk.W)
        self.z_col_var = tk.StringVar()
        self.z_col_combo = ttk.Combobox(col_frame, textvariable=self.z_col_var, width=15)
        self.z_col_combo.grid(row=1, column=1, padx=5)
        
        # Labels frame
        label_frame = ttk.LabelFrame(main_frame, text="Plot Labels", padding="5")
        label_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Plot labels
        ttk.Label(label_frame, text="Title:").grid(row=0, column=0, sticky=tk.W)
        self.title_var = tk.StringVar()
        ttk.Entry(label_frame, textvariable=self.title_var, width=40).grid(row=0, column=1, padx=(10, 0), sticky=(tk.W, tk.E))
        
        ttk.Label(label_frame, text="X Label:").grid(row=1, column=0, sticky=tk.W)
        self.xlabel_var = tk.StringVar()
        ttk.Entry(label_frame, textvariable=self.xlabel_var, width=20).grid(row=1, column=1, padx=(10, 0), sticky=tk.W)
        
        ttk.Label(label_frame, text="Y Label:").grid(row=2, column=0, sticky=tk.W)
        self.ylabel_var = tk.StringVar()
        ttk.Entry(label_frame, textvariable=self.ylabel_var, width=20).grid(row=2, column=1, padx=(10, 0), sticky=tk.W)
        
        ttk.Label(label_frame, text="Z Label:").grid(row=3, column=0, sticky=tk.W)
        self.zlabel_var = tk.StringVar()
        ttk.Entry(label_frame, textvariable=self.zlabel_var, width=20).grid(row=3, column=1, padx=(10, 0), sticky=tk.W)
        
        # Plot button
        ttk.Button(main_frame, text="Generate Contour Plot", command=self.create_plot).grid(row=5, column=0, columnspan=2, pady=10)
        
        # Status bar
        self.status_var = tk.StringVar(value="Ready")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN)
        status_bar.grid(row=6, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
        
        # Configure grid weights
        main_frame.columnconfigure(1, weight=1)
        label_frame.columnconfigure(1, weight=1)
        
        # Place app on top layer
        main_frame.focus_force()
        
    def browse_file(self):
        filetypes = (
            ("Excel files", "*.xlsx *.xls"),
            ("All files", "*.*")
        )
        
        filename = filedialog.askopenfilename(
            title="Select Excel file",
            initialdir=".",
            filetypes=filetypes
        )
        
        if filename:
            self.filepath = filename
            self.filepath_var.set(f"...{filename[-50::]}" if len(filename) > 50 else filename)
            self.status_var.set("File selected.")
            
            # Try to get sheet names
            try:
                xl_file = pd.ExcelFile(filename)
                self.sheet_combo["values"] = xl_file.sheet_names
                if xl_file.sheet_names:
                    self.sheet_var.set(xl_file.sheet_names[0])
                    
            except Exception as e:
                messagebox.showerror("Error", f"Could not read Excel file: {e}")
    
    def load_data(self, *args):
        if not self.filepath:
            messagebox.showwarning("Warning", "No Excel file loaded.")
            return
            
        try:
            self.status_var.set("Loading data...")
            self.root.update()
            
            # Load the data
            self.data = pd.read_excel(self.filepath, sheet_name=self.sheet_var.get())
            
            # Populate column dropdowns
            columns = self.data.columns.tolist()
            self.x_col_combo["values"] = columns
            self.y_col_combo["values"] = columns
            self.z_col_combo["values"] = columns
            
            # Set default values if available
            if len(columns) >= 3:
                self.x_col_var.set(columns[0])
                self.y_col_var.set(columns[1])
                self.z_col_var.set(columns[2])
            
            self.status_var.set(f"Data loaded successfully: {len(self.data)} rows, {len(columns)} columns.")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load data: {e}")
            self.status_var.set("Error loading data.")
    
    def create_plot(self):
        if self.data is None:
            messagebox.showwarning("Warning", "No data loaded.")
            return
        
        # Get selected columns
        x_col = self.x_col_var.get()
        y_col = self.y_col_var.get()
        z_col = self.z_col_var.get()
        
        if not all([x_col, y_col, z_col]):
            messagebox.showwarning("Warning", "Missing data columns.")
            return
            
        try:
            self.status_var.set("Generating plot...")
            self.root.update()
            
            # Create the contour plot
            self.contourf(
                self.data, x_col, y_col, z_col,
                self.title_var.get(),
                self.xlabel_var.get(),
                self.ylabel_var.get(),
                self.zlabel_var.get()
            )
            
            self.status_var.set("Plot generated successfully.")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create plot: {e}")
            self.status_var.set("Error generating plot.")
    
    def contourf(self, data, x_col, y_col, z_col, plot_title, x_label, y_label, z_label):
        # Validate columns
        if not all(col in data.columns for col in [x_col, y_col, z_col]):
            missing_cols = [col for col in [x_col, y_col, z_col] if col not in data.columns]
            raise ValueError(f"Missing data columns: {missing_cols}")
        
        # Remove rows with NaN values
        clean_data = data[[x_col, y_col, z_col]].dropna()
        
        if clean_data.empty:
            raise ValueError("No valid data points after removing NaN values")
        
        # Extract data
        X = clean_data[x_col].values
        Y = clean_data[y_col].values
        Z = clean_data[z_col].values
        
        # Create grid for interpolation
        x_min, x_max = X.min(), X.max()
        y_min, y_max = Y.min(), Y.max()
        
        grid_resolution = 100
        xi = np.linspace(x_min, x_max, grid_resolution)
        yi = np.linspace(y_min, y_max, grid_resolution)
        grid_x, grid_y = np.meshgrid(xi, yi)
        
        # Interpolate
        try:
            grid_z = griddata((X, Y), Z, (grid_x, grid_y), method="cubic", fill_value=np.nan)
            if np.all(np.isnan(grid_z)):
                grid_z = griddata((X, Y), Z, (grid_x, grid_y), method="linear", fill_value=np.nan)
        except Exception as e:
            raise ValueError(f"Interpolation failed: {e}")
        
        # Create the plot
        fig, ax = plt.subplots(figsize=(10, 8))
        contourf = ax.contourf(grid_x, grid_y, grid_z, levels=20, cmap="viridis", alpha=0.8)
        
        # Add colorbar
        cbar = plt.colorbar(contourf, ax=ax, shrink=0.8)
        
        # Scatter plot of actual data points
        ax.scatter(X, Y, c=Z, cmap="viridis", s=30, edgecolors="white", linewidth=0.5, alpha=0.9)
        
        # Labels and formatting
        if plot_title: ax.set_title(plot_title, fontsize=14, fontweight="bold", pad=20)
        ax.set_xlabel(x_label if x_label else x_col, fontsize=12)
        ax.set_ylabel(y_label if y_label else y_col, fontsize=12)
        cbar.set_label(z_label if z_label else z_col, rotation=270, labelpad=20)
        
        ax.grid(True, alpha=0.3)
        
        plt.tight_layout()
        plt.show()
        
        return fig, ax

def main():
    root = tk.Tk()
    ContourPlotGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
