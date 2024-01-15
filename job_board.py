from tkinter import ttk, filedialog, Tk, Frame, RIDGE,\
    BOTH, NO,YES, CENTER , W, Menu, FALSE, Label, BOTTOM, N
from tkinter.colorchooser import askcolor
from pandas import read_excel, DataFrame
from configparser import ConfigParser
from os import startfile

# create gui
root = Tk ()
root.title ("Machine Shop Job Board")

# sets state to full screen
root.state ('zoomed')
df = DataFrame
data_valid = False

# Create an object of Style widget
style = ttk.Style ()
aktualTheme = style.theme_use ()
style.theme_create ("dummy", parent=aktualTheme)
style.theme_use ("dummy")

# Create a Frame that hold the table
frame = Frame (root, relief=RIDGE)
frame.pack (anchor=N, padx=10, pady=10, fill=BOTH)

# create tree view
my_tree = ttk.Treeview (frame)

# create config file
config = ConfigParser ()

# checks & creates config file if doesnt exist, and defaults
def create_config_file():
    config.read ('config.ini')
    try:
        config.add_section ('main')
        # file data
        config.set ('main', 'file name', '')
        config.set ('main', 'sheet name', 'Schedule')
        config.set ('main', 'empty cell', '---')
        config.set ('main', 'table use status', 'True')
        config.set ('main', 'update interval', '5000')
        config.set ('main', 'default font', '11')
        config.set('main', 'max frame height', '680')
        # colors
        config.set ('main', 'my blue', '#005ed1')
        config.set ('main', 'running1', '#00ca00')
        config.set ('main', 'running2', '#349c28')
        config.set ('main', 'fp1', '#ff4a4a')
        config.set ('main', 'fp2', '#ff6c6c')
        config.set ('main', 'setup1', '#eee618')
        config.set ('main', 'setup2', '#ffff80')
        config.set ('main', 'queue1', '#9e9e98')
        config.set ('main', 'queue2', '#9e9e98')
        config.set ('main', 'maintenance', '#50e3c2')
        config.set ('main', 'clean', '#ff8040')
        config.set ('main', 'black', '#000201')
        # font size
        # also set same variables in 'select_file' to reset fonts
        reset_font ()
        # tags
        config.set ('main', 'running tag', 'RUNNING')
        config.set ('main', 'fp tag', '1ST PIECE')
        config.set ('main', 'setup tag', 'SETUP')
        config.set ('main', 'queue tag', 'QUEUE')
        config.set ('main', 'maintenance tag', 'MAINTENANCE')
        config.set ('main', 'clean tag', 'CLEAN')
    except:
        pass

def reset_font():
    config.set ('main', 'font size', config.get ('main', 'default font'))
    config.set ('main', 'max counter', '32')
    config.set ('main', 'min font', '8')
    config.set ('main', 'max font', '20')

def save_config():
    with open ('config.ini', 'w') as f:
        config.write (f)

# updates job board every xxxx ms
def update_job_board():
    save_config ()
    open_file ()
    clear_treeview ()
    fill_table ()
    root.after (config.getint ('main', 'update interval'), update_job_board)  # run itself again after x ms

# clear table
def clear_treeview():
    my_tree.delete (*my_tree.get_children ())

# opens xcel file
def open_file():
    global df, data_valid
    if config.get ('main', 'file name') != '':
        try:
            df = read_excel (config.get ('main', 'file name'), sheet_name=config.get ('main', 'sheet name'))
        except:
            print ('Cant File Sheet Name')


# calculates the width of each row for starting points per screen size
def calculate_row_widths():
    print (f'screen width: {frame.winfo_screenwidth ()}')
    # row_ratios = [.1, .1, .1, .1, .15, .15, .29 ]
    row_ratios = [.1, .08, .08, .08, .15, .15, .4]
    row_widths = []
    for ratio in row_ratios:
        row_widths.append (round (frame.winfo_screenwidth () * ratio))
    print (f'Row Widths: {row_widths}')
    return row_widths

def get_tag(counter, data):
    # if counter is even
    # running     = r combined with e or o
    # setup       = s combined with e or o
    # first pc    = f combined with e or o
    # queue       = q combined with e or o
    # maintenance = m
    # clean       = c

    tag = ""

    if (counter % 2) == 0:
        if data == config.get ('main', 'running tag'):
            return 're'
        if data == str (config.get ('main', 'fp tag')):
            return 'fe'
        if data == config.get ('main', 'setup tag'):
            return 'se'
        if data == config.get ('main', 'queue tag'):
            return 'qe'
    if (counter % 2) != 0:
        if data == config.get ('main', 'running tag'):
            return 'ro'
        if data == str (config.get ('main', 'fp tag')):
            return 'fo'
        if data == config.get ('main', 'setup tag'):
            return 'so'
        if data == config.get ('main', 'queue tag'):
            return 'qo'
    if data == config.get ('main', 'maintenance tag'):
        return 'm'
    if data == config.get ('main', 'clean tag'):
        return 'c'
    return ' '


def format_text_size():
    # sets header font & size
    style.configure ("Treeview.Heading", font=(None, config.getint ('main', 'font size') + 6))
    # sets table font & size
    style.configure ("Treeview", rowheight=config.getint ('main', 'font size') + 8,
                     font=(None, config.getint ('main', 'font size')))

# formats treeview
def format_treeview():
    if config.get ('main', 'file name') != '':
        global df
        open_file ()
        row_width = calculate_row_widths ()
        format_text_size ()
        # Add new data in Treeview widget
        tree_columns = list (df.columns)
        my_tree["column"] = tree_columns
        my_tree["show"] = "headings"

        # gets rid of the 1st column
        my_tree.column ("#0", width=0, stretch=NO)
        my_tree.heading ("#0", text="", anchor=W)

        # For Headings iterate over the columns
        for col in my_tree["column"]:
            my_tree.heading (col, text=col)
            my_tree.column (col)
        # format columns
        counter = 0
        for record in tree_columns:
            my_tree.column (record, anchor=CENTER, minwidth=80, width=row_width[counter], stretch=YES)
            my_tree.heading (record, text=record, anchor=CENTER)
            counter += 1

# changes table & font size
def alter_table_and_height(counter):
    # controls table height
    print(f'frame height({frame.winfo_height()}) &  window height({root.winfo_screenheight()})')
    # temp font size variable
    current_font = config.getint ('main', 'font size')
    # temp max frame height
    max_frame_height = config.getint('main', 'max frame height')
    # temp row height, 8 from format_text_size function
    row_height = config.getint('main', 'font size') + 8
    buffer = 20
    if frame.winfo_height() > max_frame_height:
        if config.getint('main', 'font size') > config.getint('main', 'min font'):
            config.set ('main', 'font size', f'{str (current_font - 1)}')
        if config.getint ('main', 'font size') < config.getint ('main', 'min font'):
            config.set ('main', 'font size', f"{config.getint ('main', 'min font')}")
    elif frame.winfo_height() < max_frame_height - row_height - buffer:
        if config.getint ('main', 'font size') < config.getint ('main', 'max font'):
            config.set ('main', 'font size', f'{str (current_font + 1)}')
        if config.getint ('main', 'font size') > config.getint ('main', 'max font'):
            config.set ('main', 'font size', f"{config.getint ('main', 'max font')}")

# fills out table information
def fill_table():
    if config.get ('main', 'file name'):
        # Put Data in Rows
        # For Headings iterate over the columns and fills in empty rows with empty_cell_filler
        for col in my_tree["column"]:
            my_tree.heading (col, text=col)
            my_tree.column (col)
            df[col].fillna (config.get ('main', 'empty cell'), inplace=True)

        # converts data to usable list
        df_rows = df.to_numpy ().tolist ()

        # adds values into table
        # if row has a job number, adds row to table
        counter = 0
        for record in df_rows:
            if config.getboolean ('main', 'table use status'):
                if record[4] != config.get ('main', 'empty cell') and record[4] != 'Job Number' and record[
                    3] != config.get ('main', 'empty cell'):
                    tag = get_tag (counter=counter, data=record[3])
                    my_tree.insert (parent='', index='end', iid=counter, text="", values=record, tag=tag)
                    counter += 1
            else:
                if record[4] != config.get ('main', 'empty cell') and record[4].lower () != 'job number':
                    tag = get_tag (counter=counter, data=record[3])
                    my_tree.insert (parent='', index='end', iid=counter, text="", values=record, tag=tag)
                    counter += 1

        alter_table_and_height(counter=counter)
        # formats colors of rows from tags
        my_tree.configure (height=counter)
        my_tree.tag_configure (tagname='re', background=config.get ('main', 'running1'))
        my_tree.tag_configure (tagname='fe', background=config.get ('main', 'fp1'))
        my_tree.tag_configure (tagname='se', background=config.get ('main', 'setup1'))
        my_tree.tag_configure (tagname='qe', background=config.get ('main', 'queue1'))
        my_tree.tag_configure (tagname='ro', background=config.get ('main', 'running2'))
        my_tree.tag_configure (tagname='fo', background=config.get ('main', 'fp2'))
        my_tree.tag_configure (tagname='so', background=config.get ('main', 'setup2'))
        my_tree.tag_configure (tagname='qo', background=config.get ('main', 'queue2'))
        my_tree.tag_configure (tagname='m', background=config.get ('main', 'maintenance'))
        my_tree.tag_configure (tagname='c', background=config.get ('main', 'clean'))
        my_tree.tag_configure (tagname='machine_break', background=config.get ('main', 'black'))
        format_text_size ()
        # packs table
        my_tree.pack ()

# open file command
def select_file():
    temp_file_name = ""
    try:
        temp_file_name = filedialog.askopenfilename (
            initialdir="C:/Users/",
            title='Open a file',
            filetypes=(('Excel Files', '*.xlsx'), ('All files', '*.*'))
        )
        print (f'new file: {temp_file_name}')
        # if something was selected open file, if not do nothing
        if temp_file_name != "":
            print (f'new file: {temp_file_name}')
            config.set ('main', 'file name', temp_file_name)
            reset_font ()
            clear_treeview ()
            format_treeview ()
    except:
        print ('could not load file')
        change_footer_text ('File Not Found, Please Try Again')

# creates footer restricted label
def initFooterText(text="Company Restricted"):
    global footer_label
    footer_label = Label (root, text=text, background=config.get ('main', 'my blue'),
                                      font=("Arial", 25))
    footer_label.pack (padx=5, pady=3, side=BOTTOM)

def change_footer_text(text='Company Restricted'):
    footer_label.config (text=text)

# creates the color picker gui
def color_picker(key, label):
    colors = askcolor (title=f'Color Picker - {label} Tag', color=config.get ('main', f'{key}'))
    if colors[1] is not None:
        config.set ('main', key, colors[1])
        print (f'Color for {key} is {colors[1]}')  # 0 = rgb, 1 = hex
        save_config ()

# opens current excel file
def open_excel():
    startfile (f"{config.get ('main', 'file name')}")

# opens config file
def open_config():
    startfile (f"{'config.ini'}")

# creates menu bar on top of gui
def create_drop_down():
    menubar = Menu (root)

    # creates 'FILE' drop down
    filemenu = Menu (menubar, tearoff=FALSE)
    filemenu.add_command (label="Open", command=select_file)
    filemenu.add_command (label="Exit", command=root.quit)
    menubar.add_cascade (label="File", menu=filemenu)

    # creates colors functions
    colormenu = Menu (menubar, tearoff=0)
    colormenu.add_command (label='Running Color 1', command=lambda: color_picker ('running1', 'Running'))
    colormenu.add_command (label='Running Color 2', command=lambda: color_picker ('running2', 'Running'))
    colormenu.add_command (label='First PC Color 1', command=lambda: color_picker ('fp1', 'First PC'))
    colormenu.add_command (label='First PC Color 2', command=lambda: color_picker ('fp2', 'First PC'))
    colormenu.add_command (label='Setup Color 1', command=lambda: color_picker ('setup1', 'Setup'))
    colormenu.add_command (label='Setup Color 2', command=lambda: color_picker ('setup2', 'Setup'))
    colormenu.add_command (label='Queue Color 1', command=lambda: color_picker ('queue1', 'Queue'))
    colormenu.add_command (label='Queue Color 2', command=lambda: color_picker ('queue2', 'Queue'))
    colormenu.add_command (label='Maintenance Color', command=lambda: color_picker ('maintenance', 'Maintenance'))
    colormenu.add_command (label='Clean Color', command=lambda: color_picker ('clean', 'Clean'))
    menubar.add_cascade (label='Colors', menu=colormenu)

    # open active excel sheet
    excelmenu = Menu (menubar, tearoff=0)
    excelmenu.add_command (label="Open Excel", command=open_excel)
    excelmenu.add_command (label="Settings", command=open_config)
    menubar.add_cascade (label='Excel', menu=excelmenu)

    # adds drop downs to gui
    root.config (menu=menubar)

if __name__ == "__main__":
    create_config_file ()
    root.configure (background=config.get ('main', 'my blue'))
    create_drop_down ()
    initFooterText ()
    format_treeview ()
    fill_table ()
    update_job_board ()
    root.mainloop ()
