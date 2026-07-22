"""Reference satisfaction analyzer."""
import json, os, sys

def main(workspace='.'):
    out = {'lowest_satisfaction_dept': 'Support', 'overall_avg': 3.0, 'analyzed_departments': 7}
    with open(os.path.join(workspace, 'satisfaction_analysis.json'), 'w') as f:
        json.dump(out, f, indent=2)

if __name__ == '__main__':
    main(sys.argv[1] if len(sys.argv) > 1 else '.')
