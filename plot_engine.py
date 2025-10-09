import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from scipy.interpolate import griddata


class PlotEngine:
    """Handles all data loading, processing, and plotting operations."""
    
    def __init__(self):
        # Data storage
        self.data = None
        self.filepath = None
        self.sheet_names = []
    
    def load_file(self, filepath):
        # Load an Excel file and return available sheet names.
        try:
            self.filepath = filepath
            xl_file = pd.ExcelFile(filepath)
            self.sheet_names = xl_file.sheet_names
            return self.sheet_names
        except Exception as e:
            raise ValueError(f"Could not read Excel file: {e}")
    
    def load_sheet(self, sheet_name):
        # Load data from a specific sheet and return columns.
        if not self.filepath:
            raise ValueError("No Excel file loaded.")
        
        try:
            self.data = pd.read_excel(self.filepath, sheet_name=sheet_name)
            return self.data.columns.tolist()
        except Exception as e:
            raise ValueError(f"Failed to load data: {e}")
    
    def get_data_info(self):
        # Return information about the loaded data.
        if self.data is None:
            return None
        return {
            'rows': len(self.data),
            'columns': len(self.data.columns),
            'column_names': self.data.columns.tolist()
        }
    
    def create_contour_plot(self, x_col, y_col, z_col, plot_title="", 
                           x_label="", y_label="", z_label=""):
        # Create and display contour plot
        # Validate columns
        if self.data is None:
            raise ValueError("No data loaded.")
        
        if not all(col in self.data.columns for col in [x_col, y_col, z_col]):
            missing_cols = [col for col in [x_col, y_col, z_col] 
                          if col not in self.data.columns]
            raise ValueError(f"Missing data columns: {missing_cols}")
        
        # Remove rows with NaN values
        clean_data = self.data[[x_col, y_col, z_col]].dropna()
        
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
            grid_z = griddata((X, Y), Z, (grid_x, grid_y), 
                            method="cubic", fill_value=np.nan)
            if np.all(np.isnan(grid_z)):
                grid_z = griddata((X, Y), Z, (grid_x, grid_y), 
                                method="linear", fill_value=np.nan)
        except Exception as e:
            raise ValueError(f"Interpolation failed: {e}")
        
        # Create the plot
        fig, ax = plt.subplots(figsize=(10, 8))
        contourf = ax.contourf(grid_x, grid_y, grid_z, levels=20, 
                              cmap="viridis", alpha=0.8)
        
        # Add colorbar
        cbar = plt.colorbar(contourf, ax=ax, shrink=0.8)
        
        # Labels and formatting
        if plot_title:
            ax.set_title(plot_title, fontsize=14, fontweight="bold", pad=20)
        ax.set_xlabel(x_label if x_label else x_col, fontsize=12)
        ax.set_ylabel(y_label if y_label else y_col, fontsize=12)
        cbar.set_label(z_label if z_label else z_col, rotation=270, labelpad=20)
        
        ax.grid(True, alpha=0.3)
        
        plt.tight_layout()
        plt.show()
        
        return fig, ax
    
    def create_scatter_plot(self, x_col, y_col, z_col=None, plot_title="", 
                            x_label="", y_label=""):
            # Create and display scatter plot
            # Validate columns
            if self.data is None:
                raise ValueError("No data loaded.")
            
            if not all(col in self.data.columns for col in [x_col, y_col]):
                missing_cols = [col for col in [x_col, y_col] 
                            if col not in self.data.columns]
                raise ValueError(f"Missing data columns: {missing_cols}")
            
            # Remove rows with NaN values
            clean_data = self.data[[x_col, y_col, z_col]].dropna()
            
            if clean_data.empty:
                raise ValueError("No valid data points after removing NaN values")
            
            # Extract data
            X = clean_data[x_col].values
            Y = clean_data[y_col].values
            Z = clean_data[z_col].values
            
            # Create the plot
            fig, ax = plt.subplots(figsize=(10, 8))
            ax.scatter(X, Y, alpha=0.8)
            
            # Optional error bars
            if any(Z):
                ax.errorbar(X, Y, yerr=Z, fmt="o")
            
            # Labels and formatting
            if plot_title:
                ax.set_title(plot_title, fontsize=14, fontweight="bold", pad=20)
            ax.set_xlabel(x_label if x_label else x_col, fontsize=12)
            ax.set_ylabel(y_label if y_label else y_col, fontsize=12)
            
            ax.grid(True, alpha=0.3)
            
            plt.tight_layout()
            plt.show()
            
            return fig, ax