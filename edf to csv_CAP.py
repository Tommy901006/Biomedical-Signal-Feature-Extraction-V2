import numpy as np
import scipy.io
import mne
import pandas as pd

# ---------- Step 1: 計算 CAP A 區間 ----------
def CAP(time_tot, duration, hyp):
    startB = time_tot + duration
    startA = time_tot
    durationB = np.diff(np.concatenate(([startA[0]], startB)))

    CAPs = []
    CAPf = []
    cap_a_intervals = []

    count = 0
    countCAP = 0

    for k in range(len(durationB)):
        if durationB[k] < 60:
            count += 1
        else:
            if count >= 2:
                countCAP += 1
                CAPf.append(k)
                CAPs.append(k - count)
                cap_a_intervals.append((startA[k - count], startA[k]))
            count = 0

        if k == len(durationB) - 1 and count >= 2:
            countCAP += 1
            CAPf.append(k) 
            CAPs.append(k - count + 1)
            cap_a_intervals.append((startA[k - count], startB[k]))  

    # 避免 index 超出範圍
    CAPs = np.array(CAPs)
    CAPf = np.array(CAPf)

    if len(CAPs) == 0 or len(CAPf) == 0:
        CAPtime = 0
        rate = 0
    else:
        CAPtime = np.sum(startA[CAPf] - startA[CAPs])
        NREMtime = np.sum((hyp[:, 0] != 5) & (hyp[:, 0] != 0)) * 30
        rate = CAPtime / NREMtime if NREMtime > 0 else 0

    return CAPtime, rate, cap_a_intervals


# ---------- Step 2: 讀取資料 ----------
# 載入 .mat 檔
strbrux_data = scipy.io.loadmat('No pathology\micro_strn10.mat')
hyp_data = scipy.io.loadmat('No pathology\hypn10.mat')

# 提取資料
time_tot = strbrux_data['time_tot'].flatten()
duration = strbrux_data['duration'].flatten()
hyp = hyp_data['hyp']

# 計算 CAP A 區間
CAPtime, rate, cap_a_intervals = CAP(time_tot, duration, hyp)
print(f"CAP time: {CAPtime:.2f} sec, CAP rate: {rate:.2%}, CAP intervals: {len(cap_a_intervals)}")

# ---------- Step 3: 讀 EDF 檔 ----------
edf_file = r'No pathology\n10.edf'  # 
raw = mne.io.read_raw_edf(edf_file, preload=True)
print(f"Sampling rate: {raw.info['sfreq']} Hz")
# ---------- Step 4: 擷取所有區間並合併儲存為一個 CSV ----------
all_segments = []

for idx, (start_sec, end_sec) in enumerate(cap_a_intervals):
    try:
        segment = raw.copy().crop(tmin=start_sec, tmax=end_sec)
        data, times = segment.get_data(return_times=True)
        times = times + start_sec  # 修正為 EDF 中的實際時間
        df = pd.DataFrame(data.T, columns=raw.ch_names)
        df.insert(0, 'Time(sec)', times)
        df.insert(0, 'Segment', idx + 1)  # 按照擷取順序編號
        all_segments.append(df)
        print(f"[✓] Segment {idx+1} extracted: {start_sec}–{end_sec} sec")
    except Exception as e:
        print(f"[✗] Segment {idx+1} failed: {e}")

# 合併所有段落並按 Segment 編號排序（照提取順序）
final_df = pd.concat(all_segments, ignore_index=True)
final_df = final_df.sort_values(by='Segment').reset_index(drop=True)

# 輸出為一個 CSV 檔案
final_df.to_csv("nfle11.csv", index=False)
print("[✓] all_OK")
