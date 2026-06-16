import { useRef, useState } from 'react'
import { ImagePlus, Loader2, AlertCircle, Trash2 } from 'lucide-react'
import { cn } from '../../lib/utils'
import type { OverviewImage } from '../../types/wizard'

interface Props {
  value: OverviewImage | null
  onChange: (image: OverviewImage | null) => void
}

const MAX_FILE_SIZE = 10 * 1024 * 1024 // 10MB original upload
const ACCEPTED_TYPES = ['image/jpeg', 'image/png', 'image/webp', 'image/gif']
// Target matches the OVERVIEW sheet's F3:K14 image anchor (~693x320px, ~2.16:1).
const TARGET_W = 1400
const TARGET_H = 648
const JPEG_QUALITY = 0.85

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

// Center-crops the source image to the F3:K14 target aspect ratio, downscales to
// TARGET_W x TARGET_H, and re-encodes as JPEG so the embedded file stays small and
// never gets stretched/distorted by ExcelJS's cell-range image anchor.
async function cropAndResize(file: File): Promise<string> {
  const dataUrl = await readFileAsDataURL(file)
  const img = await loadImage(dataUrl)

  const targetRatio = TARGET_W / TARGET_H
  const srcW = img.naturalWidth
  const srcH = img.naturalHeight
  const srcRatio = srcW / srcH

  let sx = 0
  let sy = 0
  let sw = srcW
  let sh = srcH
  if (srcRatio > targetRatio) {
    sw = Math.round(srcH * targetRatio)
    sx = Math.round((srcW - sw) / 2)
  } else if (srcRatio < targetRatio) {
    sh = Math.round(srcW / targetRatio)
    sy = Math.round((srcH - sh) / 2)
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
      const dataUrl = await cropAndResize(file)
      onChange({ dataUrl })
    } catch {
      setError('Could not process this image. Please try a different file.')
    } finally {
      setIsProcessing(false)
    }
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
    return (
      <div className="relative">
        <img
          src={value.dataUrl}
          alt="Cover preview"
          className="w-full aspect-[2.16/1] object-cover rounded-xl border border-gray-200"
        />
        <button
          type="button"
          onClick={() => onChange(null)}
          className="absolute top-2 right-2 bg-white/90 rounded-full p-1.5 shadow-sm hover:bg-red-50 hover:text-red-500 transition-colors"
          aria-label="Remove cover photo"
        >
          <Trash2 size={15} />
        </button>
        <p className="mt-1.5 text-xs text-gray-400">
          This image will appear at the top of the OVERVIEW sheet.
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
