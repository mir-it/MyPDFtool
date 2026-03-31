import sys
from pathlib import Path

import classify_pages as cp


def classify(path: Path):
    s, o = cp.classify_pdf(path)
    return {"S": set(s), "O": set(o), "S_list": s, "O_list": o}


def main():
    sys.stdout.reconfigure(encoding="utf-8")

    _here = Path(__file__).resolve().parent
    _data = _here / "data"
    _e = "\u042d\u043a\u0437"
    files = {
        "Ex1": _data / f"{_e}.pdf",
        "Ex2": _data / f"{_e}2.pdf",
        "Ex3": Path(f"C:/Users/wdevi/Downloads/{_e}3.pdf"),
        "Ex4": Path(f"C:/Users/wdevi/Downloads/{_e}4.pdf"),
        "Ex5": Path(f"C:/Users/wdevi/Downloads/{_e} 5.pdf"),
    }
    cur = {k: classify(p) for k, p in files.items()}

    # Approved expectations from chat.
    exp1_s = set(
        [
            2,
            3,
            4,
            5,
            6,
            7,
            8,
            9,
            11,
            12,
            13,
            14,
            15,
            17,
            18,
            21,
            22,
            23,
            24,
            26,
            27,
            28,
            29,
            30,
            31,
            32,
            33,
            35,
            36,
            40,
            42,
            43,
            44,
            46,
            47,
            50,
            51,
            52,
            53,
            55,
            56,
            57,
            62,
            63,
            64,
            66,
            67,
            69,
            70,
            71,
            72,
            75,
            76,
            77,
            78,
            79,
            80,
            81,
            82,
            83,
            89,
            90,
            91,
            92,
            93,
            95,
            97,
            98,
            99,
            100,
        ]
    )
    exp1_o = set([1, 10, 16, 19, 20, 25, 34, 37, 38, 39, 41, 45, 48, 49, 54, 58, 59, 60, 61, 65, 68, 73, 74, 84, 85, 86, 87, 88, 94, 96])

    baseline_s2 = set([1, 5, 6, 8, 9, 11, 12, 22, 23, 27, 29, 33, 34, 36, 37, 38, 43, 44, 45, 48, 49, 51, 54, 55, 56, 59, 60, 66, 69, 74, 76, 77])
    baseline_s2 |= set([7, 13, 15, 17, 21, 30, 47, 52, 63, 64, 67, 73, 75, 79, 80])
    baseline_s2.discard(49)
    exp2_s_first80 = baseline_s2

    exp3_s = set([1, 4, 5, 6, 7, 9, 13, 16, 18, 19, 20, 24, 25, 26, 29, 30, 32, 34, 35, 39])
    exp3_o = set([2, 3, 8, 10, 11, 12, 14, 15, 17, 21, 22, 23, 27, 28, 31, 33, 36, 37, 38, 40, 41])

    exp4_points = {1: "O", 14: "S"}
    exp5_points = {1: "S", 2: "S", 3: "S", 4: "S", 16: "S", 22: "S", 30: "S"}

    print("CURRENT COUNTS:")
    for k in files:
        print(f"- {k}: S={len(cur[k]['S'])} O={len(cur[k]['O'])}")

    print("\nDIFFS VS APPROVALS:")
    d1 = sorted(cur["Ex1"]["S"] ^ exp1_s)
    print(f"- Ex1 full mismatches: {len(d1)}")
    if d1:
        print(" ", d1)

    m2 = []
    for p in range(1, 81):
        got = "S" if p in cur["Ex2"]["S"] else "O"
        want = "S" if p in exp2_s_first80 else "O"
        if got != want:
            m2.append((p, want, got))
    print(f"- Ex2 first80 mismatches: {len(m2)}")
    if m2:
        print(" ", m2[:40])

    d3 = sorted(cur["Ex3"]["S"] ^ exp3_s)
    print(f"- Ex3 full mismatches: {len(d3)}")
    if d3:
        print(" ", d3)

    m4 = []
    for p, w in exp4_points.items():
        g = "S" if p in cur["Ex4"]["S"] else "O"
        if g != w:
            m4.append((p, w, g))
    print(f"- Ex4 checkpoints mismatches: {len(m4)} {m4}")

    m5 = []
    for p, w in exp5_points.items():
        g = "S" if p in cur["Ex5"]["S"] else "O"
        if g != w:
            m5.append((p, w, g))
    print(f"- Ex5 checkpoints mismatches: {len(m5)} {m5}")


if __name__ == "__main__":
    main()

