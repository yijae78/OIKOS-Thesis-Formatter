"""Save translation batches to JSON files."""
import json
import sys
import os

def save_batch(batch_file, translations_dict):
    """Save a translation batch to JSON."""
    outdir = os.path.join(os.path.dirname(__file__), '..', 'output')
    path = os.path.join(outdir, batch_file)
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(translations_dict, f, ensure_ascii=False, indent=2)
    print(f"Saved {len(translations_dict)} translations to {batch_file}")

def merge_all_batches(output_file='translations_all.json'):
    """Merge all batch files into one."""
    outdir = os.path.join(os.path.dirname(__file__), '..', 'output')
    merged = {}
    for fname in sorted(os.listdir(outdir)):
        if fname.startswith('translations_') and fname.endswith('.json') and fname != output_file:
            path = os.path.join(outdir, fname)
            with open(path, 'r', encoding='utf-8') as f:
                batch = json.load(f)
                merged.update(batch)
    path = os.path.join(outdir, output_file)
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(merged, f, ensure_ascii=False, indent=2)
    print(f"Merged {len(merged)} translations to {output_file}")
    return merged

if __name__ == '__main__':
    if len(sys.argv) > 1 and sys.argv[1] == 'merge':
        merge_all_batches()
