# import os
# from pathlib import Path
# import torch, ffmpeg, numpy as np
# from transformers import AutoModelForSpeechSeq2Seq, AutoProcessor, pipeline

# # 固定路径 ── 按需修改
# INPUT_FOLDER  = Path(r"C:\Users\13684\Desktop\jophin\joplindirectoutput\_resources")   # .m4a 所在目录
# OUTPUT_FOLDER = Path(r"C:\Users\13684\Desktop\jophin\joplin_audio2txt")  # .txt 输出目录

# # ─────────────── 加载语音识别模型 ───────────────
# def load_transcription_pipeline():
#     print("正在加载 Whisper-large-v3-turbo（首次运行需下载）…")
#     device      = "cuda:0" if torch.cuda.is_available() else "cpu"
#     torch_dtype = torch.float16 if torch.cuda.is_available() else torch.float32
#     model_id    = "openai/whisper-large-v3-turbo"

#     model = AutoModelForSpeechSeq2Seq.from_pretrained(
#         model_id,
#         torch_dtype=torch_dtype,
#         low_cpu_mem_usage=True,
#         use_safetensors=True,
#     ).to(device)

#     processor = AutoProcessor.from_pretrained(model_id)

#     asr = pipeline(
#         "automatic-speech-recognition",
#         model=model,
#         tokenizer=processor.tokenizer,
#         feature_extractor=processor.feature_extractor,
#         chunk_length_s=30,
#         batch_size=32,
#         return_timestamps=False,
#         torch_dtype=torch_dtype,
#         device=device,
#     )
#     return asr

# # ─────────────── 音频转写函数 ───────────────
# def transcribe_audio(m4a_path: Path, asr_pipe) -> str:
#     with open(m4a_path, "rb") as f:
#         m4a_bytes = f.read()

#     out, _ = (
#         ffmpeg
#         .input("pipe:0")
#         .output("pipe:1", format="s16le", acodec="pcm_s16le", ac=1, ar=16000)
#         .run(input=m4a_bytes, capture_stdout=True, capture_stderr=True, quiet=True)
#     )
#     audio_np = np.frombuffer(out, np.int16).astype(np.float32) / 32768.0
#     result   = asr_pipe(audio_np, generate_kwargs={"language": "chinese"})
#     return result["text"].strip() or "[空结果]"

# # ─────────────── 主流程 ───────────────
# def main():
#     if not INPUT_FOLDER.exists():
#         print(f"[致命] 输入目录不存在：{INPUT_FOLDER}")
#         return
#     OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)

#     m4a_files = list(INPUT_FOLDER.glob("*.m4a"))
#     if not m4a_files:
#         print(f"[提示] 在 {INPUT_FOLDER} 未找到 .m4a 文件。")
#         return

#     asr_pipe = load_transcription_pipeline()

#     for audio_fp in m4a_files:
#         txt_path = OUTPUT_FOLDER / f"{audio_fp.stem}.txt"

#         # ① 先检查目标 .txt 是否已存在 —— 若存在直接跳过
#         if txt_path.exists():
#             print(f"跳过已存在的转录文件：{txt_path.name}")
#             continue

#         try:
#             print(f"➜ 正在转写 {audio_fp.name} …")
#             transcript = transcribe_audio(audio_fp, asr_pipe)
#             txt_path.write_text(transcript, encoding="utf-8")
#             print(f"   ✓ 已生成 {txt_path.name}")
#         except Exception as e:
#             print(f" ✗ 处理 {audio_fp.name} 失败：{e}")
#             import traceback
#             traceback.print_exc()  # 添加这行来显示完整错误堆栈

#     print("\n==== 全部处理完毕 ====")

# if __name__ == "__main__":
#     main()



import os
from pathlib import Path
import torch, ffmpeg, numpy as np
from transformers import pipeline

# 固定路径 ── 按需修改
INPUT_FOLDER  = Path(r"C:\Users\13684\Desktop\jophin\joplindirectoutput\_resources")   # .m4a 所在目录
OUTPUT_FOLDER = Path(r"C:\Users\13684\Desktop\jophin\joplin_audio2txt")  # .txt 输出目录

# ─────────────── 加载语音识别模型 (已更新为 Belle-Whisper) ───────────────
def load_transcription_pipeline():
    """
    加载并初始化 Belle-Whisper 语音识别模型 Pipeline.
    """
    print("正在加载 Belle-whisper-large-v3-turbo（首次运行需下载）…")
    device      = "cuda:0" if torch.cuda.is_available() else "cpu"
    torch_dtype = torch.float16 if torch.cuda.is_available() else torch.float32
    model_id    = "BELLE-2/Belle-whisper-large-v3-turbo-zh"

    # 直接使用 pipeline 函数加载模型，更简洁
    transcriber = pipeline(
        "automatic-speech-recognition",
        model=model_id,
        torch_dtype=torch_dtype,
        device=device,
    )

    # 强制指定解码语言为中文，任务为转录
    transcriber.model.config.forced_decoder_ids = (
        transcriber.tokenizer.get_decoder_prompt_ids(
            language="zh",
            task="transcribe"
        )
    )
    return transcriber

# ─────────────── 音频转写函数 (已更新) ───────────────
def transcribe_audio(m4a_path: Path, asr_pipe) -> str:
    """
    读取 m4a 文件，转为所需格式，然后进行语音转写.
    """
    with open(m4a_path, "rb") as f:
        m4a_bytes = f.read()

    # 使用 ffmpeg 将 m4a 音频转为模型所需的 16kHz 单声道 PCM 格式
    out, _ = (
        ffmpeg
        .input("pipe:0")
        .output("pipe:1", format="s16le", acodec="pcm_s16le", ac=1, ar=16000)
        .run(input=m4a_bytes, capture_stdout=True, capture_stderr=True, quiet=True)
    )
    
    # 将 PCM 数据转为 float32 numpy 数组
    audio_np = np.frombuffer(out, np.int16).astype(np.float32) / 32768.0
    
    # 调用 pipeline 进行转写 (不再需要 generate_kwargs)
    result   = asr_pipe(audio_np)
    
    return result["text"].strip() or "[空结果]"

# ─────────────── 主流程 (无需修改) ───────────────
def main():
    if not INPUT_FOLDER.exists():
        print(f"[致命] 输入目录不存在：{INPUT_FOLDER}")
        return
    OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)

    m4a_files = list(INPUT_FOLDER.glob("*.m4a"))
    if not m4a_files:
        print(f"[提示] 在 {INPUT_FOLDER} 未找到 .m4a 文件。")
        return

    asr_pipe = load_transcription_pipeline()

    for audio_fp in m4a_files:
        txt_path = OUTPUT_FOLDER / f"{audio_fp.stem}.txt"

        # ① 先检查目标 .txt 是否已存在 —— 若存在直接跳过
        if txt_path.exists():
            print(f"跳过已存在的转录文件：{txt_path.name}")
            continue

        try:
            print(f"➜ 正在转写 {audio_fp.name} …")
            transcript = transcribe_audio(audio_fp, asr_pipe)
            txt_path.write_text(transcript, encoding="utf-8")
            print(f"   ✓ 已生成 {txt_path.name}")
        except Exception as e:
            print(f" ✗ 处理 {audio_fp.name} 失败：{e}")
            import traceback
            traceback.print_exc()  # 添加这行来显示完整错误堆栈

    print("\n==== 全部处理完毕 ====")

if __name__ == "__main__":
    main()
