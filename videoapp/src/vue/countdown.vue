<template>
  <label for="seconds">Be back in seconds:</label>
  <input v-model="inputSeconds" id="seconds" style="display: block" />
  <label for="inputValue">Input your message:</label>
  <input v-model="inputValue" id="inputValue" style="display: block" />
  <div
    id="html2canvas"
    style="padding: 12px 16px; background-color: #fffa; border-radius: 4px"
  >
    <p
      v-if="inputValue"
      style="text-align: center; text-shadow: 1px 1px 2px rgba(0, 0, 0, 0.2)"
    >
      {{ inputValue }}
    </p>
    <p style="text-align: center; text-shadow: 1px 1px 2px rgba(0, 0, 0, 0.2)">
      Be back in:
      <b style="font-size: 1.4rem; font-weight: bold">{{ formattedTime }}</b>
    </p>
  </div>
</template>

<script lang="ts">
import { defineComponent, PropType } from 'vue';

export default defineComponent({
  name: 'CountDown',
  props: {
    countDown: {
      type: Number as PropType<number>,
      required: true,
    },
  },
  data() {
    return {
      currentSeconds: this.countDown,
      inputSeconds: 0,
      inputValue: '',
      interval: null,
    };
  },
  mounted() {
    this.startCountdown();
  },
  unmounted() {
    this.currentSeconds = 0;
    clearInterval(this.interval);
  },
  watch: {
    inputSeconds(newVal) {
      this.currentSeconds = newVal;
    },
  },
  computed: {
    formattedTime(): string {
      const minutes = Math.floor(this.currentSeconds / 60);
      const seconds = this.currentSeconds % 60;
      return `${String(minutes).padStart(2, '0')}:${String(seconds).padStart(
        2,
        '0'
      )}`;
    },
  },
  methods: {
    startCountdown() {
      this.interval = setInterval(() => {
        if (this.currentSeconds > 0) {
          this.currentSeconds--;
        } else {
          clearInterval(this.interval);
        }
      }, 1000);
    },
  },
});
</script>

<style scoped>
h1 {
  color: blue;
}
</style>
