"""Compute advising analytics for Global Governance courses 20 and 21."""
import statistics

def analyze(scores):
    avg = sum(scores) / len(scores)
    median = statistics.median(scores)
    pass_rate = sum(1 for s in scores if s >= 60) / len(scores) * 100
    return avg, median, pass_rate

def grade_distribution(scores):
    bins = {"F": 0, "D": 0, "C": 0, "B": 0, "A": 0}
    for s in scores:
        if s < 60: bins["F"] += 1
        elif s < 70: bins["D"] += 1
        elif s < 80: bins["C"] += 1
        elif s < 90: bins["B"] += 1
        else: bins["A"] += 1
    return bins
