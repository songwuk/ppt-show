<script setup lang="ts">
// import PptxGenJS from 'PptxGenJS' // runtime
import PptxGenJS from 'pptxgenjs' // browser time
import tinycolor from 'tinycolor2' // https://github.com/bgrins/TinyColor
import * as htmlToImage from 'html-to-image'
const canvasLife = ref<HTMLElement | null>(null)
// https://support.microsoft.com/en-us/office/change-the-size-of-your-slides-040a811c-be43-40b9-8d04-0de5ed79987e#OfficeVersion=Newer_versions
const listType = ref([
  {
    type: 'LAYOUT_4x3',
    text: '4x3',
    width: 1000,
    height: 750,
  },
  {
    type: 'LAYOUT_16x9',
    text: '16x9',
    width: 1000,
    height: 562.5,
  },
  {
    type: 'LAYOUT_16x10',
    text: '16x10',
    width: 1000,
    height: 625,
  },
])
const chooseType = ref(1)
const pptCanvasWH = reactive({
  width: 0,
  height: 0,
})
watchEffect(() => {
  pptCanvasWH.width = listType.value[chooseType.value].width
  pptCanvasWH.height = listType.value[chooseType.value].height
})
const imgOnLine = ref<string>('')
fetch('https://picsum.photos/500').then(response => response.blob()).then((blob) => {
  imgOnLine.value = window.URL.createObjectURL(blob)
})
async function createImage(dataUrl: string) {
  return new Promise((resolve, reject) => {
    const img = new Image()
    img.onload = () => resolve(img)
    img.onerror = reject
    img.crossOrigin = 'anonymous'
    img.decoding = 'sync'
    img.src = dataUrl
  }) as Promise<Record<string, any>>
}
async function htmlToImg() {
  const dataUrl = await htmlToImage.toPng(canvasLife.value!)
  const mathRandom = Math.random().toString(36)
  const alink = document.createElement('a')
  const { src } = await createImage(dataUrl) as HTMLImageElement
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
// children []
const lifeChildren: Array<HTMLElement> = []
function recursionFn(life: HTMLElement) {
  lifeChildren.push(life)
  const nodeList = Array.from(life!.children) as Array<HTMLElement>
  if (life.hasChildNodes()) {
    for (const parg of nodeList) {
      if (parg.hasChildNodes())
        recursionFn(parg)
      else
        lifeChildren.push(parg)
    }
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
  const life = canvasLife.value! as HTMLElement
  recursionFn(life)
  const slide = pptx.addSlide() // 创建幻灯片
  slide.addNotes('pptist') // 添加幻灯片备注;
  const listtree = {} as Record<string, any>
  let index = -1
  for (const ele of lifeChildren) {
    index++
    const tagName = ele.tagName.toLowerCase()
    listtree[tagName + index] = {
      tagName,
      cssStyle: ele,
    }
  }
  for (const key in listtree) {
    const ele = window.getComputedStyle(listtree[key].cssStyle) // 获取样式
    const prePos = canvasLife.value!.getBoundingClientRect() as any
    const position = listtree[key].cssStyle.getBoundingClientRect() as any
    if (listtree[key].cssStyle.getAttribute('data-source') === 'bg') { // 背景图片
      const ccolor = ele.getPropertyValue('background-color')
      const c = formatColor(ccolor)
      slide.background = { color: c.color, transparency: (1 - c.alpha) * 100 }
    }
    else if (listtree[key].cssStyle.getAttribute('data-source') === 'image') { // 图片
      const el = listtree[key].cssStyle.getAttribute('src')
      const options: PptxGenJS.ImageProps = {
        path: el,
        x: (position.left - prePos.left) / 100,
        y: (position.top - prePos.top) / 100,
        w: position.width / 100,
        h: position.height / 100,
      }
      options.hyperlink = { url: 'https://github.com/songwuk/ppt-show', tooltip: 'ppt-show' } // 超链接
      const opacity = ele.getPropertyValue('opacity') as any
      if (opacity && Number(opacity) !== 0)
        options.transparency = 100 - opacity * 100
      slide.addImage(options)
    }
    else if (listtree[key].cssStyle.getAttribute('data-source') === 'shapes') { // 图形
      // ShapeType.rect
      const fillColor = formatColor(ele.getPropertyValue('border-color') as any)
      const borderStyle = ele.getPropertyValue('border-width') as any
      const bgColor = formatColor(ele.getPropertyValue('background-color') as any)
      const opacity = ele.getPropertyValue('opacity') as any === undefined ? 1 : ele.getPropertyValue('opacity') as any
      const points = formatPonits([
        { x: 0, y: 0 },
        { x: position.width / 100, y: 0 },
        { x: position.width / 100, y: position.height / 100 },
        { x: 0, y: position.height / 100 },
        { x: 0, y: 0 },
      ])
      const options: PptxGenJS.ShapeProps = {
        x: (position.left - prePos.left) / 100,
        y: (position.top - prePos.top) / 100,
        w: position.width / 100,
        h: position.height / 100,
        line: { color: fillColor.color, width: borderStyle.slice(0, -2) },
        fill: { color: bgColor.color, transparency: (1 - bgColor.alpha * opacity) * 100 },
        points,
      }
      slide.addShape('custGeom' as PptxGenJS.ShapeType, options)
    }
    else if (listtree[key].cssStyle.getAttribute('data-source') === 'text') { // 文字
      const textProps = listtree[key].cssStyle.textContent
      const defaltFontSize = ele.getPropertyValue('font-size') as any
      const options: PptxGenJS.TextPropsOptions = {
        x: (position.left - prePos.left - (+defaltFontSize.slice(0, -2) / 2 + 4)) / 100,
        y: (position.top - prePos.top) / 100,
        w: (position.width + +defaltFontSize.slice(0, -2) + 8) / 100,
        h: position.height / 100,
        fontFace: '微软雅黑',
        fontSize: 20 * 0.75,
        color: '#000000',
        valign: 'middle',
        align: 'center',
        isTextBox: false,
        margin: 0 * 0.75,
        paraSpaceBefore: 0 * 0.75,
        lineSpacingMultiple: 0 / 1.25,
        paraSpaceAfter: 0,
        autoFit: true,
        charSpacing: 1, // 字符间距
      }
      const opacity = ele.getPropertyValue('opacity') as any
      if (opacity && Number(opacity) !== 0)
        options.transparency = (1 - opacity) * 100
      const defaltColor = ele.getPropertyValue('color') as any
      if (defaltColor)
        options.color = formatColor(defaltColor).color
      if (defaltFontSize)
        options.fontSize = defaltFontSize.slice(0, -2) * 0.75
      // const defaultFontName = ele.getPropertyValue('font-family') as any
      // if (defaultFontName)
      //   options.fontFace = defaultFontName.split(',')[0]
      const letterSpacing = ele.getPropertyValue('letter-spacing') as any
      // const padding = ele.getPropertyValue('padding') as any
      // if (padding) {
      //   const paddingArr = padding.split(' ')
      //   if (paddingArr.length === 1) {
      //     options.margin = paddingArr[0] * 0.75
      //   }
      //   else if (paddingArr.length === 2) {
      //     options.margin = paddingArr[0] * 0.75
      //     options.paraSpaceBefore = paddingArr[1] * 0.75
      //   }
      //   else if (paddingArr.length === 3) {
      //     options.margin = paddingArr[0] * 0.75
      //     options.paraSpaceBefore = paddingArr[1] * 0.75
      //     options.paraSpaceAfter = paddingArr[2] * 0.75
      //   }
      //   else if (paddingArr.length === 4) {
      //     options.margin = paddingArr[0] * 0.75
      //     options.paraSpaceBefore = paddingArr[1] * 0.75
      //     options.paraSpaceAfter = paddingArr[2] * 0.75
      //     options.lineSpacingMultiple = paddingArr[3] / 1.25
      //   }
      // }
      letterSpacing === 'normal' && letterSpacing
        ? options.charSpacing = 1 * 0.75
        : options.charSpacing = letterSpacing.slice(0, -2) * 0.75
      slide.addText(textProps, options)
    }
  }
  pptx.writeFile()
  lifeChildren.splice(0, lifeChildren.length)
}
function checkButton(idx: number) {
  chooseType.value = idx
}
</script>

<template>
  <div px20 dark:bg-black flex="~ col" justify-center items-center>
    <div flex="~ gap2" justify-center items-center my2>
      <button btn text-sm @click="htmlToImg">
        下载图片
      </button>
      <button btn text-sm @click="htmlToPptx">
        下载PPTX
      </button>
    </div>
    <div
      ref="canvasLife" bg-gray-200 b-gray-300 b-width-2 data-source="bg" :style="{
        width: `${pptCanvasWH.width}px`,
        height: `${pptCanvasWH.height}px`,
      }"
    >
      <button data-source="text" block>
        测试文字1
      </button>
      <img v-show="imgOnLine" w100 opacity60 ma :src="imgOnLine" alt="图片" mb5 mt5 data-source="image">
      <div mb1 data-source="shapes" b-width-2 b-red-300 bg-green-100 flex="~" justify-center items-center>
        <p color-black text-2xl data-source="text" opacity60 color-red style="letter-spacing:4px">
          测试文字
        </p>
      </div>
      <button btn text-sm data-source="shapes">
        <p data-source="text">
          测试按钮1
        </p>
      </button>
      <button btn mb5 text-sm data-source="shapes" block>
        <p data-source="text">
          测试按钮2
        </p>
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
