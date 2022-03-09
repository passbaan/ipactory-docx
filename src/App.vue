<template>
  <main>
    <Test />
  </main>
</template>
<script>
import Test from "@/components/Test.vue";
export default {
  name: "AppComponent",
  components: {
    Test,
  },
};
</script>

<style lang="scss">
@import "./assets/app.vue";
</style>
