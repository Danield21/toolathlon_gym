"""Research synthesizer script - reads metadata and content, produces synthesis."""
import json
import os


def main():
    here = os.path.dirname(os.path.abspath(__file__))
    with open(os.path.join(here, "papers_metadata.json")) as f:
        meta = json.load(f)
    with open(os.path.join(here, "paper_contents.json")) as f:
        content = json.load(f)

    methods = []
    citation_network = {}
    relevance_scores = {}

    for paper in meta.get("papers", []):
        pid = paper.get("id")
        relevance_scores[pid] = paper.get("citation_count", 0) / 100.0
        for ref in paper.get("references", []):
            citation_network.setdefault(pid, []).append(ref)

    for c in content.get("contents", []):
        for m in c.get("key_methods", []):
            methods.append({"paper_id": c.get("id"), "method": m})

    synthesis = {
        "key_methods": methods,
        "citation_network": citation_network,
        "relevance_scores": relevance_scores,
    }
    with open(os.path.join(here, "research_synthesis.json"), "w") as f:
        json.dump(synthesis, f, indent=2)
    print("Wrote research_synthesis.json")


if __name__ == "__main__":
    main()
