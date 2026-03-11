"""
Aegis Mod - AI エンジン (β)
- ローカル: sentence-transformers / 感情分析
- 外部API: OpenAI
- 学習データCSV化・フィードバック蓄積
- 酷すぎるコメントスコアリング
"""
import os, json, csv, time, threading
from datetime import datetime
from difflib import SequenceMatcher

# ─────────────────────────────────────────
#  定数
# ─────────────────────────────────────────
AEGISMOD_DIR = os.path.join(os.path.expanduser("~"), ".aegismod")
os.makedirs(AEGISMOD_DIR, exist_ok=True)
AI_DATA_FILE = os.path.join(AEGISMOD_DIR, "ai_data.json")
AI_CSV_FILE  = os.path.join(AEGISMOD_DIR, "training.csv")

# 日本語の暴言・誹謗中傷ワードセット（ローカル簡易判定用）
JP_TOXIC_WORDS = {
    "死ね","殺す","殺せ","うざい","きもい","消えろ","消えて","氏ね","しね",
    "ゴミ","クズ","馬鹿","バカ","アホ","マジキチ","キチガイ","頭おかしい",
    "最悪","うせろ","失せろ","来るな","邪魔","不要","いらない","下手くそ",
    "下手糞","雑魚","ざこ","ノロマ","のろま","チート","悔しい","嫌い",
    "大嫌い","ブス","ブサイク","デブ","ハゲ","臭い","くさい","最低",
    "終わり","終わってる","無能","無駄","役立たず","カス","負け犬",
}

# スコア重み
SCORE_EXACT_REPEAT  = 15   # 完全一致連投
SCORE_SIM_REPEAT    = 12   # 類似連投
SCORE_SPEED         = 10   # 速度違反
SCORE_TOXIC_WORD    = 20   # 毒性ワード1個につき
SCORE_CAPS_RATE     = 8    # 大文字多用
SCORE_SYMBOL_SPAM   = 10   # 記号スパム
SCORE_OPENAI_BOOST  = 30   # OpenAI毒性判定ブースト

# ─────────────────────────────────────────
#  ライブラリ遅延インポート
# ─────────────────────────────────────────
_st_model       = None   # sentence-transformers
_st_loaded      = False
_openai_client  = None

def _load_st_model():
    global _st_model, _st_loaded
    if _st_loaded: return _st_model
    try:
        from sentence_transformers import SentenceTransformer
        import numpy as np
        _st_model = SentenceTransformer("paraphrase-multilingual-MiniLM-L12-v2")
        _st_loaded = True
    except Exception:
        _st_model  = None
        _st_loaded = True
    return _st_model

def _get_openai(api_key):
    global _openai_client
    if _openai_client: return _openai_client
    try:
        import openai
        _openai_client = openai.OpenAI(api_key=api_key)
    except Exception:
        _openai_client = None
    return _openai_client

# ─────────────────────────────────────────
#  AIデータストア
# ─────────────────────────────────────────
class AIDataStore:
    """学習データ・フィードバックの永続化"""
    def __init__(self):
        self.data = self._load()

    def _load(self):
        if os.path.exists(AI_DATA_FILE):
            try:
                with open(AI_DATA_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass
        return {"feedback": [], "session_scores": {}}

    def save(self):
        try:
            with open(AI_DATA_FILE, "w", encoding="utf-8") as f:
                json.dump(self.data, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def add_feedback(self, uid, msg, score, kind, correct):
        """正解/不正解フィードバックを追加"""
        self.data["feedback"].append({
            "ts":      datetime.now().isoformat(),
            "uid":     uid,
            "msg":     msg,
            "score":   score,
            "kind":    kind,
            "correct": correct,
        })
        self.save()

    def export_csv(self, path=None):
        """学習データCSVエクスポート"""
        out = path or AI_CSV_FILE
        rows = self.data.get("feedback", [])
        if not rows: return False
        with open(out, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=["ts","uid","msg","score","kind","correct"])
            w.writeheader(); w.writerows(rows)
        return out

    def get_stats(self):
        fb = self.data.get("feedback", [])
        total    = len(fb)
        correct  = sum(1 for r in fb if r.get("correct"))
        accuracy = (correct / total * 100) if total else 0
        return {"total": total, "correct": correct, "accuracy": round(accuracy, 1)}

# ─────────────────────────────────────────
#  スコアリングエンジン
# ─────────────────────────────────────────
class ScoreEngine:
    """コメントの毒性スコアを計算"""

    def __init__(self):
        self.store       = AIDataStore()
        self._score_hist = {}   # uid -> list of (ts, score, msg)
        self._lock       = threading.Lock()

    # ── ローカル簡易スコア ──
    def _local_score(self, msg: str) -> dict:
        score  = 0
        detail = []

        # 毒性ワード
        msg_l = msg.lower()
        hit_words = [w for w in JP_TOXIC_WORDS if w in msg_l]
        if hit_words:
            s = min(len(hit_words) * SCORE_TOXIC_WORD, 60)
            score += s
            detail.append(f"毒性ワード({len(hit_words)}個)+{s}")

        # 大文字多用（英語コメント）
        alpha = [c for c in msg if c.isalpha()]
        if len(alpha) >= 6:
            caps_rate = sum(1 for c in alpha if c.isupper()) / len(alpha)
            if caps_rate >= 0.7:
                score += SCORE_CAPS_RATE
                detail.append(f"大文字多用+{SCORE_CAPS_RATE}")

        # 記号スパム
        sym = sum(1 for c in msg if c in "!?！？#$%&@*")
        if sym >= 5:
            score += SCORE_SYMBOL_SPAM
            detail.append(f"記号スパム+{SCORE_SYMBOL_SPAM}")

        # 繰り返し文字（例: wwwww, aaaaa）
        import re
        if re.search(r"(.)\1{4,}", msg):
            score += 8
            detail.append("繰り返し文字+8")

        return {"score": min(score, 100), "detail": detail, "method": "local"}

    # ── OpenAI スコア ──
    def _openai_score(self, msg: str, api_key: str) -> dict:
        try:
            client = _get_openai(api_key)
            if not client:
                return None
            resp = client.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[{
                    "role": "system",
                    "content": (
                        "あなたはTwitchチャットのモデレーターです。"
                        "以下のコメントの毒性スコアを0〜100の整数で評価してください。"
                        "0=全く問題なし、100=極めて有害。"
                        "JSONで {\"score\": 数値, \"reason\": \"理由\"} のみ返してください。"
                    )
                }, {
                    "role": "user",
                    "content": f'コメント: "{msg}"'
                }],
                max_tokens=80,
                temperature=0,
            )
            raw = resp.choices[0].message.content.strip()
            # JSON パース
            import re
            m = re.search(r'\{.*\}', raw, re.DOTALL)
            if m:
                d = json.loads(m.group())
                return {
                    "score":  min(int(d.get("score", 0)), 100),
                    "detail": [d.get("reason", "")],
                    "method": "openai",
                }
        except Exception as e:
            return {"score": 0, "detail": [f"OpenAI Error: {e}"], "method": "openai_error"}
        return None

    # ── Embedding類似度（sentence-transformers）──
    def embedding_similarity(self, a: str, b: str) -> float:
        """意味的類似度 0〜100 を返す"""
        model = _load_st_model()
        if not model:
            # フォールバック: difflib
            return int(SequenceMatcher(None, a.lower(), b.lower()).ratio() * 100)
        try:
            import numpy as np
            embs = model.encode([a, b])
            cos  = float(np.dot(embs[0], embs[1]) /
                         (np.linalg.norm(embs[0]) * np.linalg.norm(embs[1]) + 1e-9))
            return int(cos * 100)
        except Exception:
            return int(SequenceMatcher(None, a.lower(), b.lower()).ratio() * 100)

    # ── メインスコア計算 ──
    def calc_score(self, uid: str, msg: str, kind: str,
                   api_key: str = "", use_openai: bool = False) -> dict:
        """
        kind: "exact" | "similar" | "speed" | "ng" | "chat"（通常チャット）
        returns: {"score": 0-100, "detail": [...], "method": str}
        """
        result = self._local_score(msg)

        # 違反種別ボーナス
        kind_bonus = {
            "exact":   SCORE_EXACT_REPEAT,
            "similar": SCORE_SIM_REPEAT,
            "speed":   SCORE_SPEED,
        }.get(kind, 0)
        if kind_bonus:
            result["score"] = min(result["score"] + kind_bonus, 100)
            result["detail"].append(f"{kind}違反+{kind_bonus}")

        # OpenAI で上書き
        if use_openai and api_key:
            oa = self._openai_score(msg, api_key)
            if oa and oa["method"] == "openai":
                # ローカルスコアとブレンド
                blended = int(result["score"] * 0.3 + oa["score"] * 0.7)
                result["score"]  = blended
                result["detail"] += oa["detail"]
                result["method"] = "openai+local"

        # 履歴に追加
        with self._lock:
            if uid not in self._score_hist:
                self._score_hist[uid] = []
            self._score_hist[uid].append({
                "ts":    time.time(),
                "score": result["score"],
                "msg":   msg,
                "kind":  kind,
            })
            # 直近100件のみ保持
            self._score_hist[uid] = self._score_hist[uid][-100:]

        return result

    # ── ユーザー別集計 ──
    def get_user_stats(self, uid: str) -> dict:
        hist = self._score_hist.get(uid, [])
        if not hist:
            return {"avg": 0, "max": 0, "count": 0, "total_score": 0}
        scores = [h["score"] for h in hist]
        return {
            "avg":         round(sum(scores) / len(scores), 1),
            "max":         max(scores),
            "count":       len(hist),
            "total_score": sum(scores),
        }

    def get_ranking(self, top_n=10) -> list:
        """酷すぎるコメントランキング（全ユーザー合算スコア順）"""
        rows = []
        with self._lock:
            for uid, hist in self._score_hist.items():
                if not hist: continue
                scores = [h["score"] for h in hist]
                rows.append({
                    "uid":         uid,
                    "total_score": sum(scores),
                    "avg_score":   round(sum(scores)/len(scores), 1),
                    "max_score":   max(scores),
                    "count":       len(hist),
                    "worst_msg":   max(hist, key=lambda x: x["score"])["msg"],
                })
        rows.sort(key=lambda x: x["total_score"], reverse=True)
        return rows[:top_n]

    def get_mvp(self) -> dict | None:
        """最強酷すぎるコメンテーター賞"""
        ranking = self.get_ranking(1)
        return ranking[0] if ranking else None

    def add_feedback(self, uid, msg, score, kind, correct: bool):
        self.store.add_feedback(uid, msg, score, kind, correct)

    def export_training_csv(self, path=None):
        return self.store.export_csv(path)

    def get_feedback_stats(self):
        return self.store.get_stats()

    def clear_session(self):
        with self._lock:
            self._score_hist.clear()

# ─────────────────────────────────────────
#  シングルトン
# ─────────────────────────────────────────
_engine = None
def get_engine() -> ScoreEngine:
    global _engine
    if _engine is None:
        _engine = ScoreEngine()
    return _engine
