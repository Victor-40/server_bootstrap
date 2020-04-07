from flask import Flask, jsonify, request
from flask_cors import CORS
import json
import os
import re
import sqlite3
import vix
import subprocess
import sys
from win32com import client
import pythoncom

# configuration
DEBUG = True

# instantiate the app
app = Flask(__name__)
app.config.from_object(__name__)

# enable CORS
CORS(app, resources={r'/api/*': {'origins': '*'}})
# CORS(app)

root_nv = r'\\svr-rum-net-04\new_versions'
root_host_test = r'D:\Testing\Test-1'
root_guest_test = r'c:\Test'
root_report = r'\\rum-cherezov-dt\!Reports'
# cfg_path = r'c:\production_svelte\server\cfg.json'
# snap_cfg_path = r'c:\production_svelte\server\snap_dct.json'
db_path = r'c:\production_svelte\server\db.sqlite3'
snapshot_dct = dict()
all_cfg_dct = dict()
prod_cfg_dct = dict()

# ---- DB ----
conn = sqlite3.connect(db_path)
cursor = conn.cursor()

res = cursor.execute("SELECT vm_name, vm_path, vm_snap, lang, prod_prefix, production, cad FROM fenix_maindb")
all_recs = res.fetchall()


# res = cursor.execute("SELECT vm_name, vm_path, vm_snap, prod_prefix  FROM fenix_maindb WHERE production='1'")
# all_snapshots = res.fetchall()
cursor = conn.cursor()
full_prod = cursor.execute("SELECT prod_root FROM prod_dirs").fetchall()

for vm, path, snap, lang, prefix, production, cad in all_recs:
    # print(vm, path, snap, lang, prefix, production, cad)
    if vm in all_cfg_dct:

        all_cfg_dct[vm]['snap'].append(snap)
    else:
        all_cfg_dct[vm] = {'path': path, 'lang': lang, 'snap': [snap]}





################################
# cursor = conn.cursor()
#
# res = cursor.execute("SELECT vm_name,  vm_snap, lang  FROM fenix_maindb").fetchall()

# for item in res:
#     # print(item)
#     vm, snap, lang = item
#     # print(vm, snap)
#     if vm in all_cfg_dct:
#         all_cfg_dct[vm]['snap'].append(snap)
#     else:
#         all_cfg_dct[vm] = {'snap': [snap], 'lang': lang}
#         # pprint.pprint(my_dict)

for item in all_cfg_dct:
    all_cfg_dct[item]['snap'] = sorted(all_cfg_dct[item]['snap'])

conn.close()
################ ---- end DB ----

# for item in snapshot_row:
#     vm, snap = item
#     if vm in snapshot_dct:
#         snapshot_dct[vm].append(snap)
#     else:
#         snapshot_dct[vm] = [snap]
#
# for key in snapshot_dct:
#     snapshot_dct[key] = sorted(snapshot_dct[key])


host = vix.VixHost(service_provider=3)


def find_builds(build, tag, _prod, subdir):

    patt = re.compile(r'-%s(_x64)*__(git--)*%s$' % (build, tag), re.I)

    work_prod = list()
    for i in _prod:
        for j in full_prod:
            prefix, = j
            if prefix.startswith(i):
                work_prod.append(prefix)

    search_dirs = [os.path.join(root_nv, item, subdir) for item in work_prod]
    setups = list()

    for _dir in search_dirs:
        if os.path.exists(_dir):
            obj = os.scandir(_dir)
            for item in obj:
                if re.search(patt, item.name):
                    setups.append(item.path)
    return setups


def make_xls(setups):
    result = list()

    # with open(snap_cfg_path) as fi:
    #     vms = json.load(fi)

    for _setup in setups:
        setup_prefix = os.path.basename(_setup).split('-')[0]
#          vm_name, vm_path, vm_snap, lang, prod_prefix, production, cad
        for vm_name, vm_path, vm_snap, _, prod_prefix, production, _ in all_recs:

            if prod_prefix.startswith(setup_prefix) and production == "1":
                result.append((_setup,  vm_name, vm_path,  vm_snap, "0"))

    job_file = r'd:\Testing\VMWare\VM-Monitor.Jobs.xls'
    if os.path.exists(job_file):
        os.remove(job_file)
    pythoncom.CoInitialize()
    xls = client.Dispatch("Excel.Application")

    wb = xls.Workbooks.Add()
    sheet = wb.WorkSheets("Sheet1")
    sheet.Name = "Table"

    # header
    header_list = ["InstallPath", "Name", "Path", "SnapName", "Done"]
    for i in range(len(header_list)):
        sheet.Cells(1, i + 1).value = header_list[i]

    for i in range(len(result)):
        for j in range(5):
            sheet.Cells(i + 2, j + 1).value = result[i][j]

    wb.SaveAs(job_file, 56)
    wb.Close()
    pythoncom.CoUninitialize()

    return result


# sanity check route
@app.route('/ping', methods=['GET'])
def ping_pong():
    return jsonify('pong!')


@app.route('/api/cfg', methods=['GET'])
def all_books():
    cfg = dict()
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    res = cursor.execute("SELECT vm_name, vm_path FROM fenix_maindb GROUP BY vm_name")
    db_req = res.fetchall()
    conn.close()
    for vmname, vmpath in db_req:
        cfg[vmname] = {'pth': vmpath}

    for _vm in cfg:
        try:
            vm = host.open_vm(cfg[_vm]['pth'])
        except vix.VixError as e:
            print(e)
            print(cfg[_vm]['pth'])
            sys.exit(1)
        if vm.is_running:
            cfg[_vm]['status'] = 'busy'
        else:
            cfg[_vm]['status'] = 'free'

    response_object = cfg
    # print(cfg)
    return jsonify(response_object)


@app.route('/api/findsetups', methods=['GET', 'POST'])
def find_setups():
    # response_object = []
    if request.method == 'POST':
        post_data = request.get_json()
        # print(post_data)
        response_object = find_builds(post_data['build'], post_data['tag'], post_data['products'], post_data['subdir'])
    else:
        response_object = ['get']

    return jsonify(response_object)


@app.route('/api/makexls', methods=['POST'])
def makexls():
    # response_object = []
    post_data = request.get_json()
    response_object = make_xls(post_data)

    return jsonify(response_object)


@app.route('/api/startclear', methods=['GET'])
def start_clear():
    # print(all_cfg_dct)
    return jsonify(all_cfg_dct)


@app.route('/api/allcfg', methods=['GET'])
def all_cfg():
    # print(snapshot_dct)
    return jsonify(all_cfg_dct)


@app.route('/api/start_testset', methods=['GET'])
def start_testset():
    subprocess.call([r'd:\Testing\VMWare\start_auto.bat'])
    return jsonify('OK! Testset started')


if __name__ == '__main__':
    app.run()
