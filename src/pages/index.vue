<script setup lang="ts">
import pptxgen from 'pptxgenjs'
import * as htmlToImage from 'html-to-image'
const canvasLife = ref<HTMLElement | null>(null)
const listType = ref([
  {
    type: 'LAYOUT_4x3',
    text: '4x3',
  },
  {
    type: 'LAYOUT_16x9',
    text: '16x9',
  },
  {
    type: 'LAYOUT_16x10',
    text: '16x10',
  },
  {
    type: 'LAYOUT_WIDE',
    text: 'wide',
  },
])
const chooseType = ref(0)
async function htmlToImg() {
  // eslint-disable-next-line new-cap
  const pptx = new pptxgen()
  const dataUrl = await htmlToImage.toPng(canvasLife.value!)
  pptx.layout = listType.value[chooseType.value].type as string // 定义比例
  pptx.defineSlideMaster({ // 定义幻灯片母版
    title: 'PPTIST_MASTER',
    background: { color: '35A2CD' },
  })
  const slide = pptx.addSlide({ masterName: 'PPTIST_MASTER' }) // 创建幻灯片
  const el = canvasLife.value!.getBoundingClientRect() as any
  slide.addImage({
    path: dataUrl,
    x: el.left / 100,
    y: el.top / 100,
    w: el.width / 100,
    h: el.height / 100,
  })
  // const slide2 = pptx.addSlide()
  // slide2.addImage({
  //   path: dataUrl,
  //   w: '100%',
  // })
  pptx.writeFile()
}
function htmlToPptx() {
  const elements = Array.from(document.querySelectorAll('*')).filter(
    (el) => {
      const eleName = el.tagName.toLowerCase()
      return eleName !== 'head' && eleName !== 'meta'
      && eleName !== 'link'
      && eleName !== 'script'
      && eleName !== 'style'
      && eleName !== 'title'
    },
  )
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
  // for (const list in listtree) {

  // }
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
    <div ref="canvasLife" bg-gray-200 dark:bg-gray-700>
      <img ma src="https://picsum.photos/500" alt="图片" mb5 py5>
      <div flex="~ row" justify-center items-center>
        <p text-sm dark:color="black">
          比例:
        </p>
        <button v-for="item, index in listType" :key="index" :style="{ background: index === chooseType ? 'rgba(13,148,136,1)' : '' }" m3 btn bg-gray-400 @click="checkButton(index)">
          {{ item.text }}
        </button>
      </div>
    </div>
  </div>
</template>
