"""
Preprocess for scholarly-nlp-survey-word task.
Clears and injects NLP transformer papers into scholarly.scholar_papers.
"""
import os
import argparse
import json

import psycopg2

DB_CONN = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": int(os.environ.get("PGPORT", "5432")),
    "dbname": "toolathlon_gym",
    "user": "eigent",
    "password": "camel",
}

NOISE_PAPERS = [
    {
        "id": 980,
        "title": "A Survey of 3D Point Cloud Processing with Deep Learning",
        "authors": [{"name": "Zhijian Tan"}, {"name": "Hao Li"}, {"name": "Yu Zhou"}],
        "abstract": "We review recent advances in processing 3D point cloud data using deep neural networks across indoor scene understanding, autonomous driving, and object detection. Methods surveyed include PointNet, PointNet++, DGCNN, and transformer-based architectures. This survey focuses on geometric point cloud data rather than language or text data.",
        "pub_year": 2021,
        "venue": "IEEE TPAMI",
        "citation_count": 2500,
    },
    {
        "id": 981,
        "title": "Soil Microbiome Dynamics Under Climate Change: A Metagenomic Survey",
        "authors": [{"name": "Maria Rodriguez"}, {"name": "James Thornton"}, {"name": "Akiko Takeda"}],
        "abstract": "We survey metagenomic studies of soil microbiome community composition and function under climate change. Our analysis includes amplicon sequencing and shotgun metagenomics data from temperate, tropical, and arctic biomes. This work belongs to environmental microbiology and is not related to natural language processing.",
        "pub_year": 2022,
        "venue": "Nature Microbiology",
        "citation_count": 800,
    },
]

NLP_PAPERS = [
    {
        "id": 169,
        "title": "Attention Is All You Need",
        "authors": [{"name": "Ashish Vaswani"}, {"name": "Noam Shazeer"}, {"name": "Niki Parmar"}],
        "abstract": "The dominant sequence transduction models are based on complex recurrent or convolutional neural networks that include an encoder and a decoder. The best performing models also connect the encoder and decoder through an attention mechanism. We propose a new simple network architecture, the Transformer, based solely on attention mechanisms, dispensing with recurrence and convolutions entirely.",
        "pub_year": 2017,
        "venue": "NeurIPS 2017",
        "citation_count": 120000,
    },
    {
        "id": 170,
        "title": "BERT: Pre-training of Deep Bidirectional Transformers for Language Understanding",
        "authors": [{"name": "Jacob Devlin"}, {"name": "Ming-Wei Chang"}, {"name": "Kenton Lee"}],
        "abstract": "We introduce a new language representation model called BERT, which stands for Bidirectional Encoder Representations from Transformers. Unlike recent language representation models, BERT is designed to pre-train deep bidirectional representations from unlabeled text by jointly conditioning on both left and right context in all layers.",
        "pub_year": 2019,
        "venue": "NAACL 2019",
        "citation_count": 95000,
    },
    {
        "id": 171,
        "title": "RoBERTa: A Robustly Optimized BERT Pretraining Approach",
        "authors": [{"name": "Yinhan Liu"}, {"name": "Myle Ott"}, {"name": "Naman Goyal"}],
        "abstract": "Language model pretraining has led to significant performance gains but careful comparison between different approaches is challenging. We present a replication study of BERT pretraining that carefully measures the impact of many key hyperparameters and training data size. We find that BERT was significantly undertrained.",
        "pub_year": 2019,
        "venue": "arXiv preprint",
        "citation_count": 18000,
    },
    {
        "id": 172,
        "title": "Exploring the Limits of Transfer Learning with a Unified Text-to-Text Transformer",
        "authors": [{"name": "Colin Raffel"}, {"name": "Noam Shazeer"}, {"name": "Adam Roberts"}],
        "abstract": "Transfer learning, where a model is first pre-trained on a data-rich task before being fine-tuned on a downstream task, has emerged as a powerful technique in natural language processing. We introduce a unified framework that converts all text-based language problems into a text-to-text format.",
        "pub_year": 2020,
        "venue": "JMLR 2020",
        "citation_count": 22000,
    },
    {
        "id": 173,
        "title": "GPT-4 Technical Report",
        "authors": [{"name": "OpenAI"}],
        "abstract": "We report the development of GPT-4, a large-scale, multimodal model which can accept image and text inputs and produce text outputs. While less capable than humans in many real-world scenarios, GPT-4 exhibits human-level performance on various professional and academic benchmarks.",
        "pub_year": 2023,
        "venue": "arXiv preprint",
        "citation_count": 8000,
    },
    {
        "id": 174,
        "title": "Training language models to follow instructions with human feedback",
        "authors": [{"name": "Long Ouyang"}, {"name": "Jeff Wu"}, {"name": "Xu Jiang"}],
        "abstract": "Making language models bigger does not inherently make them better at following a user's intent. We show an avenue for aligning language models with user intent on a wide range of tasks by fine-tuning with human feedback. Our InstructGPT models show improvements in truthfulness and reductions in toxic output generation.",
        "pub_year": 2022,
        "venue": "NeurIPS 2022",
        "citation_count": 12000,
    },
]


def inject_papers(conn):
    """Clear and inject NLP papers plus noise papers into scholarly.scholar_papers."""
    with conn.cursor() as cur:
        cur.execute("DELETE FROM scholarly.scholar_papers")
        for p in NLP_PAPERS + NOISE_PAPERS:
            cur.execute("""
                INSERT INTO scholarly.scholar_papers (id, title, authors, abstract, pub_year, venue, citation_count)
                VALUES (%s, %s, %s, %s, %s, %s, %s)
            """, (p["id"], p["title"], json.dumps(p["authors"]),
                  p["abstract"], p["pub_year"], p["venue"], p["citation_count"]))
    conn.commit()
    print(f"[preprocess] Injected {len(NLP_PAPERS)} NLP + {len(NOISE_PAPERS)} noise papers into scholarly.scholar_papers")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", type=str, required=False)
    parser.add_argument("--launch_time", type=str, required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONN)
    try:
        inject_papers(conn)
    finally:
        conn.close()

    print("\n[preprocess] Preprocessing completed successfully!")


if __name__ == "__main__":
    main()
