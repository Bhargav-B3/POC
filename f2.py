import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from tkinter import Tk, filedialog, Button, Label, Frame, Text, Scrollbar, RIGHT, Y, END
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# Function to load and clean data
def load_and_clean_data(file_path):
    df = pd.read_excel(file_path)

    # Convert 'Order Date' column to datetime
    df['Order Date'] = pd.to_datetime(df['Order Date'])

    # Convert 'Total Revenue' to numeric, coercing errors
    df['Total Revenue'] = pd.to_numeric(df['Total Revenue'], errors='coerce')

    # Drop rows with missing data
    df = df.dropna()

    return df

# Function to perform basic analysis
def analyze_data(df):
    total_sales = df['Total Revenue'].sum()
    avg_sales = df['Total Revenue'].mean()
    sales_by_product = df.groupby('Item Type')['Total Revenue'].sum().reset_index()
    return total_sales, avg_sales, sales_by_product

# Function to generate visualizations
def generate_visualizations(df):
    fig, axes = plt.subplots(2, 2, figsize=(14, 10))
    
    # Bar Chart: Total Sales by Product
    sns.barplot(x='Item Type', y='Total Revenue', data=df.groupby('Item Type')['Total Revenue'].sum().reset_index(), ax=axes[0, 0])
    axes[0, 0].set_title('Total Sales by Product')

    # Line Chart: Sales Over Time
    df.set_index('Order Date').resample('D').sum()['Total Revenue'].plot(ax=axes[0, 1])
    axes[0, 1].set_title('Sales Over Time')
    axes[0, 1].set_ylabel('Total Revenue')

    # Pie Chart: Sales Distribution by Product
    sales_by_product = df.groupby('Item Type')['Total Revenue'].sum()
    sales_by_product.plot(kind='pie', ax=axes[1, 0], autopct='%1.1f%%', startangle=140, colors=sns.color_palette("pastel"))
    axes[1, 0].set_ylabel('')
    axes[1, 0].set_title('Sales Distribution by Product')

    # Empty plot for alignment
    axes[1, 1].axis('off')

    plt.tight_layout()
    return fig

# Function to generate the summary report
def generate_summary_report(total_sales, avg_sales, sales_by_product):
    report = f"Summary Report\n"
    report += f"===============\n"
    report += f"Total Sales: ${total_sales:,.2f}\n"
    report += f"Average Sales: ${avg_sales:,.2f}\n\n"
    report += "Sales by Product:\n"
    report += sales_by_product.to_markdown(index=False)
    return report

# GUI Functionality
def show_gui():
    def on_load():
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            df = load_and_clean_data(file_path)
            total_sales, avg_sales, sales_by_product = analyze_data(df)

            # Generate summary report
            report = generate_summary_report(total_sales, avg_sales, sales_by_product)

            # Clear the frame
            for widget in frame.winfo_children():
                widget.destroy()

            # Display the summary report in a Text widget
            text_box = Text(frame, wrap="word", height=15, width=60)
            text_box.insert(END, report)
            text_box.pack(side=RIGHT, fill='both', expand=True)

            # Add a scrollbar for the Text widget
            scrollbar = Scrollbar(frame, command=text_box.yview)
            scrollbar.pack(side=RIGHT, fill=Y)
            text_box.config(yscrollcommand=scrollbar.set)

            # Create visualizations
            fig = generate_visualizations(df)
            canvas = FigureCanvasTkAgg(fig, master=frame)
            canvas.draw()
            canvas.get_tk_widget().pack(side=RIGHT, fill='both', expand=True)

    # Create the main window
    root = Tk()
    root.title("Data Analysis Tool")

    # Create a frame to hold the widgets
    frame = Frame(root)
    frame.pack(fill='both', expand=True)

    # Create a button to load the data
    Button(root, text="Load Sales Data", command=on_load).pack(pady=10)

    # Run the main loop
    root.mainloop()

if __name__ == "__main__":
    show_gui()
