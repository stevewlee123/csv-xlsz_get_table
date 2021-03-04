from flask import Flask, render_template, request, redirect, url_for, flash
import os, shutil
app = Flask(__name__)

app.config['UPLOAD_EXTENSIONS'] = ['.xlsx', '.csv']

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/', methods=['POST'])
def upload_file():

    uploaded_file = request.files['file']
    if uploaded_file.filename != '':
        file_ext = os.path.splitext(uploaded_file.filename)[1]
        if file_ext not in app.config['UPLOAD_EXTENSIONS']:
            return "Wrong format of the file. Please refresh the page."
            # delete_folder("data")
            # flash('Looks like you have changed your name!')
        uploaded_file.save("data/"+uploaded_file.filename)
    return redirect(url_for('index'))

def helper():
    import xlrd  #引入模块
    workbook=xlrd.open_workbook("data/Asset Class - Firm Map.xlsx")  #文件路径

    asset_dataset=[]
    table = workbook.sheets()[0]
    for row in range(table.nrows):
        asset_dataset.append(table.row_values(row))

    workbook=xlrd.open_workbook("data/Firm Info.xlsx")  #文件路径

    firm_dataset=[]
    table = workbook.sheets()[0]
    for row in range(table.nrows):
        firm_dataset.append(table.row_values(row))



    firm_map={}
    firm_set=set()

    for i in range(1,len(firm_dataset)):

        fid,fname=firm_dataset[i]

        firm_set.add(fname)

        firm_map[fname]=fid

    firm_set=sorted(firm_set)


    firmid_asset_map={}

    asset_set=[]

    for i in range(1,len(asset_dataset)):

        _,asset,fid=asset_dataset[i]

        if asset not in asset_set:
            asset_set.append(asset)

        firmid_asset_map.setdefault(fid,[]).append(asset)

    asset_set=[""]+sorted(asset_set)




    result=[]

    for firm in firm_set:

        tmp=[firm]
        if firm_map[firm]  in firmid_asset_map:
            asset_list=firmid_asset_map[firm_map[firm]]
            for i in asset_set:
                if i in asset_list[1:]:
                    tmp.append("X")
                else:
                    tmp.append("")
        result.append(tmp)

    for i in range(1,len(asset_set)):
        asset_set[i]+="__"
    return asset_set,result

@app.route("/get_table/",methods=['POST'])
def move_forward():
    headings,data=helper()
    return render_template("index.html",headings=headings, data=data)


if __name__ == '__main__':
    app.run()
