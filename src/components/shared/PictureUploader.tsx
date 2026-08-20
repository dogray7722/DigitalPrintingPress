import { useRef, useState } from 'react'
import { ImagePlus, Loader2, AlertCircle, Trash2, ChevronUp, ChevronDown } from 'lucide-react'
import { cn } from '../../lib/utils'
import type { OverviewImage } from '../../types/wizard'

interface Props {
  value: OverviewImage | null
  onChange: (image: OverviewImage | null) => void
}

const MAX_FILE_SIZE = 10 * 1024 * 1024 // 10MB original upload
const ACCEPTED_TYPES = ['image/jpeg', 'image/png', 'image/webp', 'image/gif']
// Target matches the OVERVIEW sheet's G1:L11 cover-photo anchor: G:L is 99 width units
// (~693px) and rows 1–11 total 260pt (~347px), so ~2:1. A mismatch here stretches the
// embedded photo, so this and the OVERVIEW row heights for rows 1–11 must change together.
const TARGET_W = 1400
const TARGET_H = 700
const TARGET_RATIO = TARGET_W / TARGET_H
const JPEG_QUALITY = 0.85
const NUDGE_STEP = 0.1

function readFileAsDataURL(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = () => resolve(reader.result as string)
    reader.onerror = () => reject(new Error('Failed to read file'))
    reader.readAsDataURL(file)
  })
}

function loadImage(src: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const img = new Image()
    img.onload = () => resolve(img)
    img.onerror = () => reject(new Error('Failed to load image'))
    img.src = src
  })
}

// The crop only has vertical slack to nudge within when the source is taller than the
// target ratio (it's the axis trimmed to fit); wider sources trim left/right instead.
function hasVerticalSlack(srcW: number, srcH: number): boolean {
  return srcW / srcH < TARGET_RATIO
}

// Crops the source image to the F1:K14 target aspect ratio at the given vertical focal
// point, downscales to TARGET_W x TARGET_H, and re-encodes as JPEG so the embedded file
// stays small and never gets stretched/distorted by ExcelJS's cell-range image anchor.
function drawCrop(img: HTMLImageElement, offsetY: number): string {
  const srcW = img.naturalWidth
  const srcH = img.naturalHeight
  const srcRatio = srcW / srcH

  let sx = 0
  let sy = 0
  let sw = srcW
  let sh = srcH
  if (srcRatio > TARGET_RATIO) {
    sw = Math.round(srcH * TARGET_RATIO)
    sx = Math.round((srcW - sw) / 2)
  } else if (srcRatio < TARGET_RATIO) {
    sh = Math.round(srcW / TARGET_RATIO)
    sy = Math.round((srcH - sh) * offsetY)
  }

  const canvas = document.createElement('canvas')
  canvas.width = TARGET_W
  canvas.height = TARGET_H
  const ctx = canvas.getContext('2d')
  if (!ctx) throw new Error('Canvas not supported')

  // Fill white first so any source transparency doesn't turn black on JPEG re-encode.
  ctx.fillStyle = '#FFFFFF'
  ctx.fillRect(0, 0, TARGET_W, TARGET_H)
  ctx.drawImage(img, sx, sy, sw, sh, 0, 0, TARGET_W, TARGET_H)

  return canvas.toDataURL('image/jpeg', JPEG_QUALITY)
}

export function PictureUploader({ value, onChange }: Props) {
  const inputRef = useRef<HTMLInputElement>(null)
  const [isDragging, setIsDragging] = useState(false)
  const [isProcessing, setIsProcessing] = useState(false)
  const [error, setError] = useState<string | null>(null)
  // Caches the decoded source image so nudges don't need to re-read/re-decode the data URL.
  const sourceImgRef = useRef<{ src: string; img: HTMLImageElement } | null>(null)

  async function getSourceImg(sourceDataUrl: string): Promise<HTMLImageElement> {
    if (sourceImgRef.current?.src !== sourceDataUrl) {
      sourceImgRef.current = { src: sourceDataUrl, img: await loadImage(sourceDataUrl) }
    }
    return sourceImgRef.current.img
  }

  async function processFile(file: File) {
    setError(null)
    if (!ACCEPTED_TYPES.includes(file.type)) {
      setError('Please upload a JPG, PNG, WEBP, or GIF image.')
      return
    }
    if (file.size > MAX_FILE_SIZE) {
      setError('Image is too large (max 10MB).')
      return
    }
    setIsProcessing(true)
    try {
      const sourceDataUrl = await readFileAsDataURL(file)
      const img = await loadImage(sourceDataUrl)
      sourceImgRef.current = { src: sourceDataUrl, img }
      const offsetY = 0.5
      const dataUrl = drawCrop(img, offsetY)
      onChange({
        dataUrl,
        sourceDataUrl,
        sourceWidth: img.naturalWidth,
        sourceHeight: img.naturalHeight,
        offsetY,
      })
    } catch {
      setError('Could not process this image. Please try a different file.')
    } finally {
      setIsProcessing(false)
    }
  }

  async function handleNudge(direction: 'up' | 'down') {
    if (!value) return
    const delta = direction === 'up' ? -NUDGE_STEP : NUDGE_STEP
    const offsetY = Math.min(1, Math.max(0, value.offsetY + delta))
    if (offsetY === value.offsetY) return
    const img = await getSourceImg(value.sourceDataUrl)
    const dataUrl = drawCrop(img, offsetY)
    onChange({ ...value, offsetY, dataUrl })
  }

  function handleFileSelect(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0]
    if (file) processFile(file)
    e.target.value = ''
  }

  function handleDrop(e: React.DragEvent<HTMLDivElement>) {
    e.preventDefault()
    setIsDragging(false)
    const file = e.dataTransfer.files?.[0]
    if (file) processFile(file)
  }

  if (value) {
    const canReposition = hasVerticalSlack(value.sourceWidth, value.sourceHeight)
    return (
      <div className="relative">
        <img
          src={value.dataUrl}
          alt="Cover preview"
          className="w-full aspect-[1400/756] object-cover rounded-xl border border-gray-200"
        />
        <button
          type="button"
          onClick={() => onChange(null)}
          className="absolute top-2 right-2 bg-white/90 rounded-full p-1.5 shadow-sm hover:bg-red-50 hover:text-red-500 transition-colors"
          aria-label="Remove cover photo"
        >
          <Trash2 size={15} />
        </button>
        {canReposition && (
          <div className="absolute bottom-2 right-2 flex flex-col rounded-full bg-white/90 shadow-sm overflow-hidden">
            <button
              type="button"
              onClick={() => handleNudge('up')}
              disabled={value.offsetY <= 0}
              className="p-1.5 hover:bg-indigo-50 hover:text-indigo-500 transition-colors disabled:opacity-30 disabled:hover:bg-transparent disabled:hover:text-current"
              aria-label="Show more of the top of the photo"
            >
              <ChevronUp size={15} />
            </button>
            <button
              type="button"
              onClick={() => handleNudge('down')}
              disabled={value.offsetY >= 1}
              className="p-1.5 hover:bg-indigo-50 hover:text-indigo-500 transition-colors disabled:opacity-30 disabled:hover:bg-transparent disabled:hover:text-current"
              aria-label="Show more of the bottom of the photo"
            >
              <ChevronDown size={15} />
            </button>
          </div>
        )}
        <p className="mt-1.5 text-xs text-gray-400">
          This image will appear at the top of the OVERVIEW sheet.
          {canReposition && ' Use the arrows to reposition it.'}
        </p>
      </div>
    )
  }

  return (
    <div>
      <input
        ref={inputRef}
        type="file"
        accept="image/*"
        className="hidden"
        onChange={handleFileSelect}
      />
      <div
        onClick={() => inputRef.current?.click()}
        onDragOver={(e) => {
          e.preventDefault()
          setIsDragging(true)
        }}
        onDragLeave={() => setIsDragging(false)}
        onDrop={handleDrop}
        className={cn(
          'rounded-xl border-2 border-dashed p-6 text-center cursor-pointer transition-colors',
          error
            ? 'border-red-300 bg-red-50/50'
            : isDragging
              ? 'border-indigo-400 bg-indigo-50/50'
              : 'border-gray-200 hover:border-indigo-300'
        )}
      >
        {isProcessing ? (
          <div className="flex flex-col items-center gap-2 text-gray-500">
            <Loader2 size={28} className="animate-spin text-indigo-400" />
            <span className="text-sm font-medium">Processing image…</span>
          </div>
        ) : error ? (
          <div className="flex flex-col items-center gap-2 text-red-600">
            <AlertCircle size={28} />
            <span className="text-sm font-medium">{error}</span>
            <span className="text-xs text-gray-400">Click to try again</span>
          </div>
        ) : (
          <div className="flex flex-col items-center gap-2 text-gray-400">
            <ImagePlus size={28} />
            <span className="text-sm font-medium text-gray-600">Click to upload or drag & drop</span>
            <span className="text-xs">JPG, PNG, WEBP, or GIF — up to 10MB</span>
          </div>
        )}
      </div>
      <p className="mt-1.5 text-xs text-gray-400">
        This image will appear at the top of the OVERVIEW sheet.
      </p>
    </div>
  )
}
