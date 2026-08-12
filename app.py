from flask import Flask, render_template, request, send_file
from datetime import datetime
from gen import generate

app = Flask(__name__)


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/generate', methods=['POST'])
def gen():
    data = request.get_json(force=True)
    korpus = [k.strip() for k in data.get('korpus', ['К1', 'К2', 'К3']) if k.strip()]
    if not korpus:
        korpus = ['К1']

    areas = {}
    raw_areas = data.get('areas', {})
    for k in korpus:
        if k in raw_areas and isinstance(raw_areas[k], list):
            areas[k] = [float(v) for v in raw_areas[k][:7]]
        else:
            areas[k] = [1500, 2000, 2500, 4430, 4000, 5000, 1500]

    start_date = datetime.strptime(data.get('startDate', '2026-06-01'), '%Y-%m-%d')
    report_date = datetime.strptime(data.get('reportDate', '2026-06-29'), '%Y-%m-%d')
    weeks = int(data.get('weeks', 48))

    bio = generate(korpus=korpus, areas=areas,
                   start_date=start_date, report_date=report_date, weeks=weeks)

    fname = f"Отделка_SIGNAL_{len(korpus)}к_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    return send_file(bio, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True, download_name=fname)


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080, debug=True)
