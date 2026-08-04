"""疲勞狀態判斷規則。

這個模組不依賴 tkinter 或 matplotlib；調整疲勞判斷門檻與規則時，
只需要修改這裡。
"""

from dataclasses import dataclass

import pandas as pd


@dataclass(frozen=True)
class FatigueAlgorithmConfig:
    """黑線實際狀態與紅線預測狀態使用的門檻設定。"""

    fatigue_start_reaction_threshold: float = 1.6
    recovery_reaction_threshold: float = 1.6
    recovery_duration_seconds: int = 60
    init_buffer_seconds: int = 30
    eye_alert_min_threshold: float = 10
    base_recovery_threshold: float = 13
    triggered_eye_floor: float = 7
    recovery_increment: float = 3
    max_recovery_threshold: float = 28
    alpha_alert_threshold: float = 3
    no_signal_threshold: float = 0.1


DEFAULT_CONFIG = FatigueAlgorithmConfig()


def classify_actual_states(
    events_df: pd.DataFrame,
    config: FatigueAlgorithmConfig = DEFAULT_CONFIG,
) -> pd.DataFrame:
    """由反應時間建立疲勞區段（圖中的黑線）。

    第一個反應時間大於等於 fatigue_start_reaction_threshold 的 event 開始區段。
    第一個低於 recovery_reaction_threshold 的 event 開始恢復計時；其後連續
    recovery_duration_seconds 秒內所有有記錄的 event 都低於恢復門檻，才結束區段。
    恢復期間若任一 event 不再低於恢復門檻，恢復計時會重新開始。
    """
    events = events_df[["second", "react_time"]].copy()
    events["second"] = pd.to_numeric(events["second"], errors="coerce")
    events["react_time"] = pd.to_numeric(events["react_time"], errors="coerce")
    events = events.dropna(subset=["second", "react_time"]).sort_values(
        "second", kind="stable"
    )

    state_rows: list[dict[str, float | int]] = []
    fatigue_active = False
    recovery_start_second: float | None = None

    for event in events.itertuples(index=False):
        second = float(event.second)
        reaction_time = float(event.react_time)

        if not fatigue_active:
            state = int(reaction_time >= config.fatigue_start_reaction_threshold)
            fatigue_active = state == 1
            recovery_start_second = None
            state_rows.append(
                {"second": second, "react_time": reaction_time, "state_val": state}
            )
            continue

        if reaction_time >= config.recovery_reaction_threshold:
            recovery_start_second = None
            state_rows.append(
                {"second": second, "react_time": reaction_time, "state_val": 1}
            )
            continue

        if recovery_start_second is None:
            recovery_start_second = second
        recovery_end_second = recovery_start_second + config.recovery_duration_seconds

        if second < recovery_end_second:
            state_rows.append(
                {"second": second, "react_time": reaction_time, "state_val": 1}
            )
            continue

        # 若確認恢復的 event 晚於結束秒數，補一筆資料讓黑線準時切回 Alert。
        if second > recovery_end_second:
            state_rows.append(
                {
                    "second": recovery_end_second,
                    "react_time": float("nan"),
                    "state_val": 0,
                }
            )
        fatigue_active = False
        recovery_start_second = None
        state_rows.append(
            {"second": second, "react_time": reaction_time, "state_val": 0}
        )

    return pd.DataFrame(state_rows, columns=["second", "react_time", "state_val"])


def predict_fatigue_states(
    master_df: pd.DataFrame,
    config: FatigueAlgorithmConfig = DEFAULT_CONFIG,
) -> pd.DataFrame:
    """由 Alpha 與眼動累積值建立預測狀態（圖中的紅線）。

    回傳原始逐秒資料，並附加 pred_state、eye_alarm、alpha_alarm 三欄。
    """
    pred_states: list[int] = []
    eye_alarms: list[int] = []
    alpha_alarms: list[int] = []

    max_eye_sum = 0.0
    previous_eye_sum = 0.0
    triggered_eye_sum: float | None = None
    monitoring_enabled = True

    for row in master_df.itertuples(index=False):
        current_second = row.second
        alpha_sum = row.alpha_sum
        eye_sum = row.eye_sum

        if current_second < config.init_buffer_seconds:
            pred_states.append(0)
            eye_alarms.append(0)
            alpha_alarms.append(0)
            previous_eye_sum = eye_sum
            max_eye_sum = eye_sum
            continue

        if not monitoring_enabled:
            recovery_limit = config.base_recovery_threshold
            if triggered_eye_sum is not None:
                if (
                    triggered_eye_sum > config.triggered_eye_floor
                    and eye_sum <= config.triggered_eye_floor
                ):
                    triggered_eye_sum = config.triggered_eye_floor
                recovery_limit = min(
                    config.max_recovery_threshold,
                    max(
                        config.base_recovery_threshold,
                        triggered_eye_sum + config.recovery_increment,
                    ),
                )

            if eye_sum >= recovery_limit and eye_sum > previous_eye_sum:
                monitoring_enabled = True
                triggered_eye_sum = None

        if (
            monitoring_enabled
            and eye_sum > previous_eye_sum
            and eye_sum >= config.base_recovery_threshold
        ):
            max_eye_sum = eye_sum

        eye_alarm = False
        alpha_alarm = False
        if monitoring_enabled:
            eye_alarm = eye_sum <= previous_eye_sum and eye_sum > 0 if max_eye_sum > 0 else False
            alpha_alarm = alpha_sum >= config.alpha_alert_threshold

            if alpha_sum <= config.no_signal_threshold and eye_sum <= config.no_signal_threshold:
                state = 1
            elif (
                (eye_alarm and alpha_alarm)
                or (eye_sum <= config.eye_alert_min_threshold and eye_sum != max_eye_sum)
            ):
                state = 1
                monitoring_enabled = False
                triggered_eye_sum = eye_sum
            else:
                state = 0
        else:
            state = 1

        pred_states.append(state)
        eye_alarms.append(int(eye_alarm))
        alpha_alarms.append(int(alpha_alarm))
        previous_eye_sum = eye_sum

    result = master_df.copy()
    result["pred_state"] = pred_states
    result["eye_alarm"] = eye_alarms
    result["alpha_alarm"] = alpha_alarms
    return result
