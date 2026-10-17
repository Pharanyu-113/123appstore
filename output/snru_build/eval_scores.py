# Expert efficiency evaluation scores (5 experts x 12 items) as given by the author on 2026-10-17.
# Expert 2 listed 3.3=5 twice; the duplicate is counted once.
import statistics as st, json, io
items = ["1.1","1.2","1.3","1.4","2.1","2.2","2.3","2.4","3.1","3.2","3.3","3.4"]
scores = {
    1: [5,5,3,3, 4,5,5,5, 3,4,5,4],
    2: [4,4,5,5, 5,5,4,5, 4,5,5,5],
    3: [3,3,3,4, 3,3,3,3, 2,4,4,3],
    4: [4,4,4,4, 5,5,4,4, 4,3,4,4],
    5: [5,5,4,4, 4,5,5,5, 3,3,5,5],
}
def level(m):
    return ("มากที่สุด" if m >= 4.21 else "มาก" if m >= 3.41 else "ปานกลาง" if m >= 2.61
            else "น้อย" if m >= 1.81 else "น้อยที่สุด")
def stat(xs): return round(st.mean(xs), 3), round(st.stdev(xs), 3)
out = {"items": {}, "sections": {}}
for i, it in enumerate(items):
    out["items"][it] = stat([scores[e][i] for e in scores])
for sec in "123":
    xs = [scores[e][i] for e in scores for i, it in enumerate(items) if it.startswith(sec + ".")]
    out["sections"][sec] = stat(xs)
out["overall"] = stat([x for e in scores for x in scores[e]])
for k, (m, s) in list(out["items"].items()) + [("sec" + k, v) for k, v in out["sections"].items()] + [("overall", out["overall"])]:
    print(f"{k:8} {m:.3f} {s:.3f} {level(m)}")
io.open("eval_stats.json", "w", encoding="utf-8").write(json.dumps(out))
