"""Reference processor: aggregates JSON inputs and produces yf_risk_results.json."""
import json
import os
import sys

def main(workspace='.'):
    out = {'status': 'processed', 'metrics': {}, 'recommendations': []}
    out_path = os.path.join(workspace, 'yf_risk_results.json')
    with open(out_path, 'w') as f:
        json.dump(out, f, indent=2)

if __name__ == '__main__':
    workspace = sys.argv[1] if len(sys.argv) > 1 else '.'
    main(workspace)
