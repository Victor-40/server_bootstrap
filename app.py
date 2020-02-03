from flask import Flask, jsonify, request
from flask_cors import CORS
import json
import os
import re
import vix
import sys

# configuration
DEBUG = True

# instantiate the app
app = Flask(__name__)
app.config.from_object(__name__)

# enable CORS
CORS(app, resources={r'/api/*': {'origins': '*'}})
# CORS(app)

BOOKS = [
    {
        'title': 'On the Road',
        'author': 'Jack Kerouac',
        'read': True
    },
    {
        'title': 'Harry Potter and the Philosopher\'s Stone',
        'author': 'J. K. Rowling',
        'read': False
    },
    {
        'title': 'Green Eggs and Ham',
        'author': 'Dr. Seuss',
        'read': True
    }
]


root_nv = r'\\svr-rum-net-04\new_versions'
root_host_test = r'D:\Testing\Test-1'
root_guest_test = r'c:\Test'
root_report = r'\\rum-cherezov-dt\!Reports'

host = vix.VixHost(service_provider=3)


def find_builds(build, tag, _prod, subdir):
    #TODO: remove absolute path
    with open(r'C:\exp_vue_bootstrap\server\cfg.json') as fi:
        cfg_dct = json.load(fi)

    patt = re.compile(r'-%s(_x64)*__(git--)*%s$' % (build, tag), re.I)

    full_prod = cfg_dct['prod_dirs']
    print(full_prod)
    work_prod = list()
    for i in _prod:
        for j in full_prod:
            if j.startswith(i):
                work_prod.append(j)

    r_dir = cfg_dct['root_dir']

    search_dirs = [os.path.join(r_dir, item, subdir) for item in work_prod]
    setups = list()

    for _dir in search_dirs:
        if os.path.exists(_dir):
            obj = os.scandir(_dir)
            for item in obj:
                if re.search(patt, item.name):
                    setups.append(item.path)
    return setups

# sanity check route
@app.route('/ping', methods=['GET'])
def ping_pong():
    return jsonify('pong!')


@app.route('/api/cfg', methods=['GET', 'POST'])
def all_books():
    response_object = {'status': 'success'}
    if request.method == 'POST':
        post_data = request.get_json()
        BOOKS.append({
            'title': post_data.get('title'),
            'author': post_data.get('author'),
            'read': post_data.get('read')
        })
        response_object['message'] = 'Book added!'
    else:
        with open(r'c:\exp_work\flask_rest\snap_dct.json') as fi:
            cfg = json.load(fi)

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
    response_object = []
    if request.method == 'POST':
        post_data = request.get_json()
        # response_object = post_data
        # response_object['status'] = 'success'
        print(post_data)
        response_object = find_builds(post_data['build'], post_data['tag'], post_data['products'], post_data['subdir'])
    else:
        response_object = ['get']

    return jsonify(response_object)


if __name__ == '__main__':
    app.run()
