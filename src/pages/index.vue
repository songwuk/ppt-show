<script setup lang="ts">
import PptxGenJS from 'PptxGenJS'
import tinycolor from 'tinycolor2' // https://github.com/bgrins/TinyColor
import * as htmlToImage from 'html-to-image'
const canvasLife = ref<HTMLElement | null>(null)
const listType = ref([
  {
    type: 'LAYOUT_4x3',
    text: '4x3',
    viewportRatio: 0.75,
  },
  {
    type: 'LAYOUT_16x9',
    text: '16x9',
    viewportRatio: 0.5625,
  },
  {
    type: 'LAYOUT_16x10',
    text: '16x10',
    viewportRatio: 0.625,
  },
  {
    type: 'LAYOUT_WIDE',
    text: 'wide',
    viewportRatio: 1,
  },
])
const chooseType = ref(0)
async function createImage(dataUrl: string): Promise<Record<string, any>> {
  return new Promise((resolve, reject) => {
    const img = new Image()
    img.onload = () => resolve(img)
    img.onerror = reject
    img.crossOrigin = 'anonymous'
    img.decoding = 'sync'
    img.src = dataUrl
  })
}
async function htmlToImg() {
  const dataUrl = await htmlToImage.toPng(canvasLife.value!)
  const mathRandom = Math.random().toString(36)
  const alink = document.createElement('a')
  const { src } = await createImage(dataUrl)
  alink.href = src
  alink.download = `${mathRandom}.png`
  alink.click()
}

// 格式化颜色值为 透明度 + HexString，供pptxgenjs使用
const formatColor = (_color: string) => {
  const c = tinycolor(_color)
  const alpha = c.getAlpha()
  const color = alpha === 0 ? '#ffffff' : c.setAlpha(1).toHexString()
  return {
    alpha,
    color,
  }
}
function htmlToPptx() {
  // https://github.com/pipipi-pikachu/PPTist/blob/985be943cacc582ad807f36aeb675ece4279c012/src/hooks/useExport.ts#L89
  const pptx = new PptxGenJS()
  pptx.layout = listType.value[chooseType.value].type as string // 定义比例
  pptx.defineSlideMaster({ // 定义幻灯片母版
    title: 'PPTIST_MASTER',
    background: { color: '35A2CD' },
  })
  // { masterName: 'PPTIST_MASTER' }
  const slide = pptx.addSlide() // 创建幻灯片
  const elements = [canvasLife.value, ...Array.from(canvasLife.value!.children)] as any[]
  const listtree = {} as any
  let index = -1
  for (const ele of elements) {
    index++
    const tagName = ele.tagName.toLowerCase()
    listtree[tagName + index] = {
      tagName,
      cssStyle: ele,
    }
  }
  for (const key in listtree) {
    const ele = window.getComputedStyle(listtree[key].cssStyle) // dom all css
    if (listtree[key].cssStyle.getAttribute('data-source') === 'bg') {
      const ccolor = ele.getPropertyValue('background-color')
      const c = formatColor(ccolor)
      slide.background = { color: c.color, transparency: (1 - c.alpha) * 100 }
    }
    else if (listtree[key].cssStyle.getAttribute('data-source') === 'image') {
      const el = listtree[key].cssStyle.getAttribute('src')
      const prePos = canvasLife.value!.getBoundingClientRect() as any
      const position = listtree[key].cssStyle.getBoundingClientRect() as any
      const options: PptxGenJS.ImageProps = {
        path: el,
        x: (position.left - prePos.left) / 100 * listType.value[chooseType.value].viewportRatio,
        y: (position.top - prePos.top) / 100 * listType.value[chooseType.value].viewportRatio,
        w: position.width / 100,
        h: position.height / 100,
      }

      options.hyperlink = { url: 'https://github.com/songwuk/ppt-show', tooltip: 'ppt-show' } // 超链接
      const opacity = ele.getPropertyValue('opacity') as any
      if (opacity && Number(opacity) !== 0)
        options.transparency = 100 - opacity * 100
      slide.addImage(options)

      console.log(position, prePos)
    }
    else {
      continue
    }
  }

  pptx.writeFile()
}
function checkButton(idx: number) {
  chooseType.value = idx
}
</script>

<template>
  <div px20 dark:bg-black>
    <div flex="~ gap2" justify-center items-center my2>
      <button btn text-sm @click="htmlToImg">
        下载图片
      </button>
      <button btn text-sm @click="htmlToPptx">
        下载PPTX
      </button>
    </div>
    <div ref="canvasLife" bg-gray-200 b-gray-300 b-width-2 data-source="bg">
      <img opacity60 ma src="https://picsum.photos/500" alt="图片" mb5 mt5 data-source="image">
      <p color-black mb1 data-source="text">
        测试文字
      </p>
      <button btn mb5 text-sm data-source="shapes">
        测试按钮
      </button>
    </div>
    <div flex="~ row" justify-center items-center>
      <p text-sm color-black>
        比例:
      </p>
      <button v-for="item, index in listType" :key="index" :style="{ background: index === chooseType ? 'rgba(13,148,136,1)' : '' }" m3 btn bg-gray-400 @click="checkButton(index)">
        {{ item.text }}
      </button>
    </div>
  </div>
</template>
