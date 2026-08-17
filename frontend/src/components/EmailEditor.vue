<template>
  <div class="email-editor">
    <div class="toolbar">
      <el-button-group>
        <el-button size="small" @click="execCmd('bold')" title="加粗"><strong>B</strong></el-button>
        <el-button size="small" @click="execCmd('italic')" title="斜体"><em>I</em></el-button>
        <el-button size="small" @click="execCmd('underline')" title="下划线"><u>U</u></el-button>
      </el-button-group>
      <el-button-group style="margin-left: 10px">
        <el-button size="small" @click="execCmd('justifyLeft')">左对齐</el-button>
        <el-button size="small" @click="execCmd('justifyCenter')">居中</el-button>
        <el-button size="small" @click="execCmd('justifyRight')">右对齐</el-button>
      </el-button-group>
      <el-button-group style="margin-left: 10px">
        <el-button size="small" @click="execCmd('insertOrderedList')">有序列表</el-button>
        <el-button size="small" @click="execCmd('insertUnorderedList')">无序列表</el-button>
      </el-button-group>
      <el-button-group style="margin-left: 10px">
        <el-button size="small" @click="insertLink">插入链接</el-button>
        <el-button size="small" @click="insertImage">插入图片</el-button>
      </el-button-group>
      <el-button-group style="margin-left: 10px">
        <el-button size="small" @click="insertVariable('{name}')">{name}</el-button>
        <el-button size="small" @click="insertVariable('{email}')">{email}</el-button>
      </el-button-group>
      <el-select v-model="fontSize" size="small" style="width: 80px; margin-left: 10px" @change="changeFontSize">
        <el-option v-for="s in [12, 14, 16, 18, 20, 24, 28]" :key="s" :label="s + 'px'" :value="s" />
      </el-select>
      <el-color-picker v-model="fontColor" size="small" style="margin-left: 10px" @change="changeFontColor" />
    </div>
    <div
      ref="editorRef"
      class="editor-body"
      contenteditable="true"
      :style="{ minHeight: minHeight }"
      @input="handleInput"
      @blur="handleInput"
      @paste.prevent="handlePaste"
      @dragover.prevent
      @drop.prevent="handleDrop"
    />
  </div>
</template>

<script setup>
import { ref, onMounted, watch } from 'vue'
import { ElMessageBox } from 'element-plus'
import DOMPurify from 'dompurify'

const props = defineProps({
  modelValue: { type: String, default: '' },
  minHeight: { type: String, default: '300px' },
})

const emit = defineEmits(['update:modelValue'])

const editorRef = ref(null)
const fontSize = ref(14)
const fontColor = ref('#333333')

const sanitizeHtml = (value) => DOMPurify.sanitize(String(value || ''), {
  USE_PROFILES: { html: true },
  ALLOWED_URI_REGEXP: /^(?:(?:https?|mailto):|[#/])/i,
})

const execCmd = (cmd, value = null) => {
  document.execCommand(cmd, false, value)
  editorRef.value?.focus()
  handleInput()
}

const changeFontSize = (size) => {
  document.execCommand('fontSize', false, '7')
  const fontElements = editorRef.value.querySelectorAll('font[size="7"]')
  fontElements.forEach(el => {
    el.removeAttribute('size')
    el.style.fontSize = size + 'px'
  })
  handleInput()
}

const changeFontColor = (color) => {
  document.execCommand('foreColor', false, color)
  handleInput()
}

const insertLink = async () => {
  try {
    const { value } = await ElMessageBox.prompt('请输入链接地址', '插入链接', {
      inputPattern: /^https?:\/\/.+/,
      inputErrorMessage: '请输入有效的URL',
    })
    document.execCommand('createLink', false, value)
    handleInput()
  } catch (e) {
    if (e !== 'cancel') ElMessage.error('插入链接失败')
  }
}

const insertImage = async () => {
  try {
    const { value } = await ElMessageBox.prompt('请输入图片地址', '插入图片', {
      inputPattern: /^https?:\/\/.+/,
      inputErrorMessage: '请输入有效的URL',
    })
    document.execCommand('insertImage', false, value)
    handleInput()
  } catch (e) {
    if (e !== 'cancel') ElMessage.error('插入链接失败')
  }
}

const insertVariable = (varName) => {
  const selection = window.getSelection()
  if (selection.rangeCount > 0) {
    const range = selection.getRangeAt(0)
    range.deleteContents()
    const textNode = document.createTextNode(varName)
    range.insertNode(textNode)
    range.collapse(false)
    selection.removeAllRanges()
    selection.addRange(range)
  } else {
    editorRef.value?.append(document.createTextNode(varName))
  }
  handleInput()
}

const handleInput = () => {
  if (editorRef.value) {
    const sanitized = sanitizeHtml(editorRef.value.innerHTML)
    if (editorRef.value.innerHTML !== sanitized) {
      editorRef.value.innerHTML = sanitized
    }
    emit('update:modelValue', sanitized)
  }
}

const insertTransferredContent = (transfer) => {
  const sourceHtml = transfer?.getData('text/html') || ''
  if (sourceHtml) {
    document.execCommand('insertHTML', false, sanitizeHtml(sourceHtml))
  } else {
    document.execCommand('insertText', false, transfer?.getData('text/plain') || '')
  }
  handleInput()
}

const handlePaste = (event) => {
  insertTransferredContent(event.clipboardData)
}

const moveCaretToPoint = (x, y) => {
  const selection = window.getSelection()
  if (!selection) return

  let range = null
  if (document.caretRangeFromPoint) {
    range = document.caretRangeFromPoint(x, y)
  } else if (document.caretPositionFromPoint) {
    const position = document.caretPositionFromPoint(x, y)
    if (position) {
      range = document.createRange()
      range.setStart(position.offsetNode, position.offset)
      range.collapse(true)
    }
  }
  if (range && editorRef.value?.contains(range.startContainer)) {
    selection.removeAllRanges()
    selection.addRange(range)
  }
}

const handleDrop = (event) => {
  editorRef.value?.focus()
  moveCaretToPoint(event.clientX, event.clientY)
  insertTransferredContent(event.dataTransfer)
}

watch(() => props.modelValue, (newVal) => {
  const sanitized = sanitizeHtml(newVal)
  if (editorRef.value && editorRef.value.innerHTML !== sanitized) {
    editorRef.value.innerHTML = sanitized
  }
})

onMounted(() => {
  if (editorRef.value && props.modelValue) {
    const sanitized = sanitizeHtml(props.modelValue)
    editorRef.value.innerHTML = sanitized
    if (sanitized !== props.modelValue) {
      emit('update:modelValue', sanitized)
    }
  }
})
</script>

<style scoped>
.email-editor {
  border: 1px solid #dcdfe6;
  border-radius: 4px;
  overflow: hidden;
}
.toolbar {
  padding: 8px;
  background: #f5f7fa;
  border-bottom: 1px solid #dcdfe6;
  display: flex;
  align-items: center;
  flex-wrap: wrap;
  gap: 4px;
}
.editor-body {
  padding: 12px;
  outline: none;
  font-size: 14px;
  line-height: 1.6;
  overflow-y: auto;
}
.editor-body:focus {
  border-color: #409eff;
}
</style>
