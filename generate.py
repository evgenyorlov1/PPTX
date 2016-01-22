#!/usr/bin/env python

from __future__ import unicode_literals

import argparse
import os
import subprocess

OUTFILE = 'output.pptx'
DASHBOARD_SEQUENCE = [
    9,
    11,
    10,
    3,
    2,
    1,
    4,
    5,
    6,
    7,
    8,
    12,
]

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('xslx-file')
    args = parser.parse_args()

    dashboard_dirs = ['dashboard-%s' % n for n in DASHBOARD_SEQUENCE]
    for dashboard_dir in dashboard_dirs:
        subprocess.call([
            os.path.join(dashboard_dir, 'main.py'),
            '--output=%s' % os.path.join(dashboard_dir, OUTFILE),
            os.path.abspath(vars(args)['xslx-file']),
        ])
    subprocess.call(
        [os.path.join('.', 'merge.py')]+[os.path.join(dashboard_dir, OUTFILE) for dashboard_dir in dashboard_dirs],
    )
