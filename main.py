import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from plot_engine import PlotEngine

class PlotAF:
    """Handles all GUI elements and user interactions."""
    
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Plot Generator")
        self.root.option_add("*tearOff", False)
        self.root.geometry("450x570")
        self.root.resizable(True, True)
        self.root.minsize(450, 570)
        
        # Initialize the plot engine
        self.engine = PlotEngine()
        
        # Import tcl file and set theme
        style = ttk.Style(self.root)
        try:
            self.root.tk.call("source", "forest-light.tcl")
            style.theme_use("forest-light")
        except:
            pass  # Theme file not found, use default
        
        # Create GUI elements
        self.setup_gui()
    
    def setup_gui(self):
        # Main frame
        main_frame = ttk.Frame(self.root, style="Card", padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Status frame
        status_frame = ttk.LabelFrame(main_frame, text="Status", padding="5")
        status_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), 
                       pady=(0, 10))
        
        # Status display
        self.status_var = tk.StringVar(value="Ready")
        status = ttk.Label(status_frame, textvariable=self.status_var)
        status.grid(row=6, column=0, columnspan=2, sticky=(tk.W, tk.E))
        
        # File selection frame
        file_frame = ttk.LabelFrame(main_frame, text="File Selection", padding="5")
        file_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), 
                       pady=(0, 10))
        
        # File path display
        self.filepath_var = tk.StringVar(value="No file selected")
        ttk.Label(file_frame, text="Selected file:").grid(row=0, column=0, 
                                                          sticky=tk.W)
        ttk.Label(file_frame, textvariable=self.filepath_var, 
                 foreground="blue").grid(row=1, column=0, sticky=tk.W)
        
        # Browse file button
        ttk.Button(main_frame, text="Browse Excel File", 
                  command=self.browse_file).grid(row=2, column=0, columnspan=2, 
                                                pady=10)
        
        # Sheet selection
        sheet_frame = ttk.LabelFrame(main_frame, text="Sheet Selection", 
                                     padding="5")
        sheet_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), 
                        pady=(0, 10))
        
        ttk.Label(sheet_frame, text="Sheet name:").grid(row=0, column=0, 
                                                        sticky=tk.W)
        self.sheet_var = tk.StringVar(value="Sheet1")
        self.sheet_combo = ttk.Combobox(sheet_frame, textvariable=self.sheet_var, 
                                       width=20)
        self.sheet_combo.grid(row=0, column=1, padx=(10, 0))
        
        # Reloads data when sheet_var is changed
        self.sheet_var.trace_add("write", self.load_data)
        
        # Column selection frame
        col_frame = ttk.LabelFrame(main_frame, text="Column Selection", 
                                   padding="5")
        col_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), 
                      pady=(0, 10))
        
        # Column dropdowns
        ttk.Label(col_frame, text="X Column:").grid(row=0, column=0, sticky=tk.W)
        self.x_col_var = tk.StringVar()
        self.x_col_combo = ttk.Combobox(col_frame, textvariable=self.x_col_var, width=15)
        self.x_col_combo.grid(row=0, column=1, padx=5)
        
        ttk.Label(col_frame, text="Y Column:").grid(row=0, column=2, sticky=tk.W)
        self.y_col_var = tk.StringVar()
        self.y_col_combo = ttk.Combobox(col_frame, textvariable=self.y_col_var, width=15)
        self.y_col_combo.grid(row=0, column=3, padx=5)
        
        self.zlab_column = ttk.Label(col_frame, text="Z Column:")
        self.zlab_column.grid(row=1, column=0, sticky=tk.W)
        self.z_col_var = tk.StringVar()
        self.z_col_combo = ttk.Combobox(col_frame, textvariable=self.z_col_var, width=15)
        self.z_col_combo.grid(row=1, column=1, padx=5)
        
        # Labels frame
        label_frame = ttk.LabelFrame(main_frame, text="Plot Labels", padding="5")
        label_frame.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Plot labels
        ttk.Label(label_frame, text="Title:").grid(row=0, column=0, sticky=tk.W)
        self.title_var = tk.StringVar()
        ttk.Entry(label_frame, textvariable=self.title_var, width=40).grid(row=0, column=1, padx=(10, 0),
                                                                           sticky=(tk.W, tk.E))
        
        ttk.Label(label_frame, text="X Label:").grid(row=1, column=0, sticky=tk.W)
        self.xlabel_var = tk.StringVar()
        ttk.Entry(label_frame, textvariable=self.xlabel_var, width=20).grid(row=1, column=1, padx=(10, 0),
                                                                            sticky=tk.W)
        
        ttk.Label(label_frame, text="Y Label:").grid(row=2, column=0, sticky=tk.W)
        self.ylabel_var = tk.StringVar()
        ttk.Entry(label_frame, textvariable=self.ylabel_var, width=20).grid(row=2, column=1, padx=(10, 0),
                                                                            sticky=tk.W)
        
        self.zlab_label = ttk.Label(label_frame, text="Z Label:")
        self.zlab_label.grid(row=3, column=0, sticky=tk.W)
        self.zlabel_var = tk.StringVar()
        self.zlab_entry = ttk.Entry(label_frame, textvariable=self.zlabel_var, width=20)
        self.zlab_entry.grid(row=3, column=1, padx=(10, 0), sticky=tk.W)
        
        # Plot type frame
        plottype_frame = ttk.LabelFrame(main_frame, text="Plot Type", padding="5")
        plottype_frame.grid(row=6, column=0, columnspan=2, sticky=(tk.W, tk.E), 
                            pady=(0, 10))
        
        # Plot type radio buttons
        self.plottype_var = tk.StringVar(value="Contour")
        ttk.Radiobutton(plottype_frame, text="Contour", variable=self.plottype_var, value="Contour").grid(row=0,
                                                                                                          column=0)
        ttk.Radiobutton(plottype_frame, text="Scatter", variable=self.plottype_var, value="Scatter").grid(row=0,
                                                                                                          column=1)
        
        # Trace to update "z" to "σ" for scatter
        self.plottype_var.trace_add("write", self.update_zlabel)
        
        # Plot button
        ttk.Button(main_frame, text="Generate Plot", 
                  style="Accent.TButton", 
                  command=self.create_plot).grid(row=7, column=0, sticky=tk.N, columnspan=2)
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(7, weight=1)
        label_frame.columnconfigure(1, weight=1)
        file_frame.columnconfigure(0, weight=1)
        sheet_frame.columnconfigure(1, weight=1)
        col_frame.columnconfigure(1, weight=1)
        col_frame.columnconfigure(3, weight=1)
        
        # Place app on top layer
        main_frame.focus_force()
    
    def browse_file(self):
        # Handles file browsing and loading.
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
            # Update display
            display_name = (f"...{filename[-50:]}" if len(filename) > 50 
                          else filename)
            self.filepath_var.set(display_name)
            self.status_var.set("File selected.")
            
            # Load file using engine
            try:
                sheet_names = self.engine.load_file(filename)
                self.sheet_combo["values"] = sheet_names
                if sheet_names:
                    self.sheet_var.set(sheet_names[0])
            except Exception as e:
                messagebox.showerror("Error", str(e))
                self.status_var.set("Error loading file.")
    
    def load_data(self, *args):
        # Load data from selected sheet.
        if not self.engine.filepath:
            messagebox.showwarning("Warning", "No Excel file loaded.")
            return
        
        try:
            self.status_var.set("Loading data...")
            self.root.update()
            
            # Load data using engine
            columns = self.engine.load_sheet(self.sheet_var.get())
            
            # Populate column dropdowns
            self.x_col_combo["values"] = columns
            self.y_col_combo["values"] = columns
            self.z_col_combo["values"] = columns
            
            # Set default values if available
            if len(columns) >= 3:
                self.x_col_var.set(columns[0])
                self.y_col_var.set(columns[1])
                self.z_col_var.set(columns[2])
            
            # Update status
            info = self.engine.get_data_info()
            self.status_var.set(
                f"Data loaded successfully: {info['rows']} rows, "
                f"{info['columns']} columns."
            )
            
        except Exception as e:
            messagebox.showerror("Error", str(e))
            self.status_var.set("Error loading data.")
    
    def update_zlabel(self, *args):
        # Change "z" to "σ" for scatter plots
        if self.plottype_var.get() == "Scatter":
            self.zlab_column["text"] = "σ Column:"
            self.zlab_label.grid_remove()
            self.zlab_entry.grid_remove()
       
        elif self.plottype_var.get() == "Contour":
            self.zlab_column["text"] = "Z Column:"
            self.zlab_label.grid()
            self.zlab_entry.grid()
            
    
    def create_plot(self):
        # Create plot from specified columns
        x_col = self.x_col_var.get()
        y_col = self.y_col_var.get()
        z_col = self.z_col_var.get()
        plot_type = self.plottype_var.get()
        
        if not all([x_col, y_col, z_col]):
            messagebox.showwarning("Warning", "Missing data columns.")
            return
        
        try:
            self.status_var.set("Generating plot...")
            self.root.update()
            
            match plot_type:
                case "Contour":
                    self.engine.create_contour_plot(
                        x_col, y_col, z_col,
                        self.title_var.get(),
                        self.xlabel_var.get(),
                        self.ylabel_var.get(),
                        self.zlabel_var.get()
                    )
                case "Scatter":
                    self.engine.create_scatter_plot(
                        x_col, y_col, z_col,
                        self.title_var.get(),
                        self.xlabel_var.get(),
                        self.ylabel_var.get()
                    )
            
            self.status_var.set("Plot generated successfully.")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create plot: {e}")
            self.status_var.set("Error generating plot.")
    
    def contour_plot(self):
        # Create contour plot from specified columns
        x_col = self.x_col_var.get()
        y_col = self.y_col_var.get()
        z_col = self.z_col_var.get()
        
        if not all([x_col, y_col, z_col]):
            messagebox.showwarning("Warning", "Missing data columns.")
            return
        
        try:
            self.status_var.set("Generating plot...")
            self.root.update()
            
            # Create plot using engine
            self.engine.create_contour_plot(
                x_col, y_col, z_col,
                self.title_var.get(),
                self.xlabel_var.get(),
                self.ylabel_var.get(),
                self.zlabel_var.get()
            )
            
            self.status_var.set("Plot generated successfully.")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create plot: {e}")
            self.status_var.set("Error generating plot.")

def main():
    root = tk.Tk()
    PlotAF(root)
    root.mainloop()

if __name__ == "__main__":
    main()
