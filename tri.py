import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from tkinter import Tk, filedialog, Button, Label, Frame, Text, Scrollbar, RIGHT, Y, END, TOP
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
    sns.set_style("darkgrid", {"axes.facecolor": "#1f2d3d"})
    fig, axes = plt.subplots(2, 2, figsize=(14, 10))
    
    # Bar Chart: Total Sales by Product
    sns.barplot(x='Item Type', y='Total Revenue', data=df.groupby('Item Type')['Total Revenue'].sum().reset_index(), ax=axes[0, 0], palette="coolwarm")
    axes[0, 0].set_title('Total Sales by Product', color="#f8f9fa")
    axes[0, 0].set_facecolor('#1f2d3d')
    axes[0, 0].tick_params(colors='#f8f9fa')

    # Line Chart: Sales Over Time
    df.set_index('Order Date').resample('D').sum()['Total Revenue'].plot(ax=axes[0, 1], color='#f76c6c')
    axes[0, 1].set_title('Sales Over Time', color="#f8f9fa")
    axes[0, 1].set_ylabel('Total Revenue', color="#f8f9fa")
    axes[0, 1].set_facecolor('#1f2d3d')
    axes[0, 1].tick_params(colors='#f8f9fa')

    # Pie Chart: Sales Distribution by Product
    sales_by_product = df.groupby('Item Type')['Total Revenue'].sum()
    sales_by_product.plot(kind='pie', ax=axes[1, 0], autopct='%1.1f%%', startangle=140, colors=sns.color_palette("coolwarm"), textprops={'color':'#f8f9fa'})
    axes[1, 0].set_ylabel('')
    axes[1, 0].set_title('Sales Distribution by Product', color="#f8f9fa")
    axes[1, 0].set_facecolor('#1f2d3d')

    # Empty plot for alignment
    axes[1, 1].axis('off')

    plt.tight_layout()
    fig.patch.set_facecolor('#1f2d3d')
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
            text_box = Text(frame, wrap="word", height=15, width=60, bg="#2e4053", fg="#f8f9fa")
            text_box.insert(END, report)
            text_box.pack(side=RIGHT, fill='both', expand=True)

            # Add a scrollbar for the Text widget
            scrollbar = Scrollbar(frame, command=text_box.yview, bg="#2e4053")
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

    # Create a frame for the navbar title
    nav_frame = Frame(root, bg='#3a4a5f')
    nav_frame.pack(fill='x')

    # Add a title to the navbar
    title_label = Label(nav_frame, text="Sales Data Analysis", font=('Arial', 18, 'bold'), bg='#3a4a5f', fg='#f8f9fa')
    title_label.pack(pady=10)

    # Create a frame to hold the widgets
    frame = Frame(root, bg='#2e4053')
    frame.pack(fill='both', expand=True)

    # Create a button to load the data
    Button(nav_frame, text="Load Sales Data", command=on_load, bg='#f76c6c', fg='#f8f9fa').pack(pady=10, padx=10, side=RIGHT)

    # Run the main loop
    root.mainloop()

if __name__ == "__main__":
    show_gui()
