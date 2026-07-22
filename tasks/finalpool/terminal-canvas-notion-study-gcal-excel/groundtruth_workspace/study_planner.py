"""Compute weekly study hours and priority per course."""

def weekly_hours(assignments, quizzes):
    return 3 + 0.5 * assignments + 1 * quizzes

def priority(weekly):
    if weekly > 8:
        return "High"
    elif weekly >= 5:
        return "Medium"
    else:
        return "Low"

if __name__ == "__main__":
    courses = []  # populated from canvas data
    for c in courses:
        wh = weekly_hours(c['assignments'], c['quizzes'])
        c['weekly_hours'] = wh
        c['priority'] = priority(wh)
