import pandas as pd
import matplotlib.pyplot as plt # type: ignore
import seaborn as sns # type: ignore
from tkinter import Tk, filedialog, Button, Label, Frame, Text, Scrollbar, LEFT, RIGHT, Y, END
from tkinter import ttk
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg # type: ignore
import matplotlib.dates as mdates # type: ignore

# Function to load and clean data
def load_and_clean_data(file_path):
    df = pd.read_excel(file_path)

    # Convert 'Order Date' column to datetime
    df['Order Date'] = pd.to_datetime(df['Order Date'])

    # Convert 'Total Revenue' to numeric, coercing errors
    df['Total Revenue'] = pd.to_numeric(df['Total Revenue'], errors='coerce')

    # Drop rows with missing data
    df = df.dropna()

    # Convert revenue to millions
    df['Total Revenue'] = df['Total Revenue'] / 1_000_000

    return df

# Function to perform basic analysis
def analyze_data(df):
    total_sales = df['Total Revenue'].sum()
    avg_sales = df['Total Revenue'].mean()
    sales_by_product = df.groupby('Item Type')['Total Revenue'].sum().reset_index()
    return total_sales, avg_sales, sales_by_product

# Function to generate visualizations
def generate_visualizations(df, sort_order='desc'):
    sns.set_style("darkgrid", {"axes.facecolor": "#2c3e50"})

    # Bar Chart: Total Sales by Product
    bar_data = df.groupby('Item Type')['Total Revenue'].sum().reset_index()
    bar_data = bar_data.sort_values(by='Total Revenue', ascending=(sort_order == 'asc'))

    fig, ax = plt.subplots(figsize=(10, 6))
    bars = sns.barplot(x='Item Type', y='Total Revenue', data=bar_data, ax=ax, palette="viridis")
    ax.set_title('Total Sales by Product', color="#e0e0e0")
    ax.set_facecolor('#2c3e50')
    ax.tick_params(colors='#e0e0e0')

    # Adding tooltips to bar chart
    def on_hover_bar(event):
        if event.inaxes == ax:
            for bar in bars.patches:
                if bar.contains(event)[0]:
                    height = bar.get_height()
                    ax.annotate(f'{height:,.2f}M',
                                (bar.get_x() + bar.get_width() / 2, height),
                                xytext=(0, 5),
                                textcoords='offset points', 
                                ha='center', va='bottom',
                                color='#e0e0e0')
                    fig.canvas.draw_idle()
                    break
            else:
                # Remove the annotation if cursor is not on any bar
                ax.annotate("", xy=(0,0), xytext=(0,0))
                fig.canvas.draw_idle()

    fig.canvas.mpl_connect("motion_notify_event", on_hover_bar)

    # Line Chart: Sales Over Time
    line_fig, line_ax = plt.subplots(figsize=(10, 6))
    line_data = df.set_index('Order Date').resample('Y').sum()['Total Revenue']
    line_ax.plot(line_data.index, line_data.values, color='#ff6f61', marker='o')

    # Format the x-axis to show years properly
    line_ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y'))
    line_ax.xaxis.set_major_locator(mdates.YearLocator())
    line_ax.set_title('Sales Over Time', color="#e0e0e0")
    line_ax.set_ylabel('Total Revenue (Millions)', color="#e0e0e0")
    line_ax.set_xlabel('Year', color="#e0e0e0")
    line_ax.set_facecolor('#2c3e50')
    line_ax.tick_params(colors='#e0e0e0')

    # Format the y-axis values
    line_ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f'{x:,.0f}M'))

    # Annotate all points
    for x, y in zip(line_data.index, line_data.values):
        line_ax.annotate(f'{y:,.0f}M',
                         xy=(x, y),
                         xytext=(10, 10),  # Adjust as needed
                         textcoords='offset points',
                         ha='center', va='bottom',
                         color='#e0e0e0',
                         fontsize=10,
                         bbox=dict(facecolor='black', alpha=0.5))

    # Pie Chart: Sales Distribution by Product
    pie_fig, pie_ax = plt.subplots(figsize=(7, 7))
    sales_by_product = df.groupby('Item Type')['Total Revenue'].sum()
    wedges, texts, autotexts = pie_ax.pie(sales_by_product, labels=sales_by_product.index, autopct='%1.1f%%',
    startangle=140, colors=sns.color_palette("magma"),
    textprops={'color': '#e0e0e0'})
    pie_ax.set_title('Sales Distribution by Product', color="#e0e0e0")
    pie_ax.set_facecolor('#2c3e50')

    # Adding tooltips to pie chart
    def on_hover_pie(event):
        if event.inaxes == pie_ax:
            for wedge in wedges:
                if wedge.contains(event)[0]:
                    wedge.set_edgecolor('yellow')
                    wedge.set_linewidth(2)
                    pie_fig.canvas.draw_idle()
                    break
            else:
                for wedge in wedges:
                    wedge.set_edgecolor('black')
                    wedge.set_linewidth(1)
                pie_fig.canvas.draw_idle()

    pie_fig.canvas.mpl_connect("motion_notify_event", on_hover_pie)

    plt.tight_layout()
    fig.patch.set_facecolor('#2c3e50')
    line_fig.patch.set_facecolor('#2c3e50')
    pie_fig.patch.set_facecolor('#2c3e50')

    return fig, line_fig, pie_fig

# Function to generate the summary report
def generate_summary_report(total_sales, avg_sales, sales_by_product):
    report = f"Summary Report\n"
    report += f"====================\n"
    report += f"Total Sales: ${total_sales:,.2f} Million\n"
    report += f"Average Sales: ${avg_sales:,.2f} Million\n\n"
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
            text_box.delete(1.0, END)

            # Display the summary report in a Text widget
            text_box.insert(END, report)

            # Create visualizations
            fig_dict = {
                'bar_desc': generate_visualizations(df, 'desc')[0],
                'bar_asc': generate_visualizations(df, 'asc')[0],
                'line': generate_visualizations(df, 'desc')[1],
                'pie': generate_visualizations(df, 'desc')[2]
            }

            def update_canvas(fig):
                for widget in canvas_frame.winfo_children():
                    widget.destroy()
                canvas = FigureCanvasTkAgg(fig, master=canvas_frame)
                canvas.draw()
                canvas.get_tk_widget().pack(side=LEFT, fill='both', expand=True)

            def show_bar_chart_asc():
                update_canvas(fig_dict['bar_asc'])
                highlight_button(bar_asc_button)

            def show_bar_chart_desc():
                update_canvas(fig_dict['bar_desc'])
                highlight_button(bar_desc_button)

            def show_line_chart():
                update_canvas(fig_dict['line'])
                highlight_button(line_button)

            def show_pie_chart():
                update_canvas(fig_dict['pie'])
                highlight_button(pie_button)

            def highlight_button(button):
                bar_desc_button.config(bg='#e74c3c', fg='#ecf0f1')
                bar_asc_button.config(bg='#e74c3c', fg='#ecf0f1')
                line_button.config(bg='#e74c3c', fg='#ecf0f1')
                pie_button.config(bg='#e74c3c', fg='#ecf0f1')
                button.config(bg='#3498db', fg='#ecf0f1')

            # Create buttons to switch between charts
            bar_desc_button = Button(nav_frame, text="Bar Chart Desc", command=show_bar_chart_desc, bg='#e74c3c', fg='#ecf0f1')
            bar_desc_button.pack(side=LEFT, padx=5)
            bar_asc_button = Button(nav_frame, text="Bar Chart Asc", command=show_bar_chart_asc, bg='#e74c3c', fg='#ecf0f1')
            bar_asc_button.pack(side=LEFT, padx=5)
            line_button = Button(nav_frame, text="Line Chart", command=show_line_chart, bg='#e74c3c', fg='#ecf0f1')
            line_button.pack(side=LEFT, padx=5)
            pie_button = Button(nav_frame, text="Pie Chart", command=show_pie_chart, bg='#e74c3c', fg='#ecf0f1')
            pie_button.pack(side=LEFT, padx=5)

            # Show the default chart
            show_bar_chart_desc()

    # Create the main window
    root = Tk()
    root.title("Data Analysis Tool")

    # Create a frame for the navbar title
    nav_frame = Frame(root, bg='#34495e')
    nav_frame.pack(fill='x')

    # Add a title to the navbar
    title_label = Label(nav_frame, text="Sales Data Analysis", font=('Arial', 18, 'bold'), bg='#34495e', fg='#ecf0f1')
    title_label.pack(pady=10)

    # Create a frame for the buttons
    button_frame = Frame(root)
    button_frame.pack(side='left', padx=10)

    # Create a frame for the canvas
    canvas_frame = Frame(root)
    canvas_frame.pack(side='right', fill='both', expand=True)

    # Create a frame for the report text box
    report_frame = Frame(root)
    report_frame.pack(side='bottom', fill='both', expand=True)

    # Create a Text widget for the report
    text_box = Text(report_frame, wrap='word', height=10, bg='#34495e', fg='#ecf0f1', font=('Arial', 12))
    text_box.pack(side=LEFT, fill=Y, expand=True)
    scrollbar = Scrollbar(report_frame, command=text_box.yview)
    scrollbar.pack(side=RIGHT, fill=Y)
    text_box.config(yscrollcommand=scrollbar.set)

    # Automatically load data on startup
    on_load()

    root.mainloop()

if __name__ == "__main__":
    show_gui()
