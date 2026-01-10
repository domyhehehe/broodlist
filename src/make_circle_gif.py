import argparse
import io
import os


def main():
    parser = argparse.ArgumentParser(description="Make a slow rotating GIF from a circle PNG")
    parser.add_argument("image", help="Path to <PK>_circle.png")
    parser.add_argument("--out", default=None, help="Output GIF path")
    parser.add_argument("--seconds", type=float, default=8.0, help="Seconds per full rotation")
    parser.add_argument("--fps", type=int, default=24, help="Frames per second")
    parser.add_argument("--max-mb", type=float, default=15.0, help="Max output size in MB")
    args = parser.parse_args()

    try:
        from PIL import Image
    except ImportError as exc:
        raise SystemExit("Pillow is required. Install with: pip install pillow") from exc

    src = args.image
    if not os.path.exists(src):
        raise SystemExit(f"Image not found: {src}")

    base = Image.open(src).convert("RGBA")
    frames = max(1, int(args.seconds * args.fps))

    out_path = args.out or f"{os.path.splitext(src)[0]}.gif"
    duration = int(1000 / args.fps)
    max_bytes = int(args.max_mb * 1024 * 1024)

    def build_frames(step):
        seq = []
        for i in range(0, frames, step):
            angle = -360.0 * i / frames
            rotated = base.rotate(angle, resample=Image.BICUBIC, expand=False)
            seq.append(rotated)
        return seq

    def save_preview(seq):
        quantized = [
            im.convert("P", palette=Image.Palette.ADAPTIVE, colors=128) for im in seq
        ]
        buf = io.BytesIO()
        quantized[0].save(
            buf,
            format="GIF",
            save_all=True,
            append_images=quantized[1:],
            duration=duration,
            loop=0,
            disposal=2,
            optimize=True,
        )
        return buf.getvalue()

    step = 1
    data = None
    while True:
        seq = build_frames(step)
        data = save_preview(seq)
        if len(data) <= max_bytes or len(seq) <= 2:
            break
        step += 1

    with open(out_path, "wb") as f:
        f.write(data)
    print(f"Wrote: {out_path}")


if __name__ == "__main__":
    main()
