<template>
  <div class="input">
    <form @submit.prevent="submitForm">
      <label for="file">Upload an html file</label>
      <input type="file" id="file" @change="fileUploaded" required />
    </form>
    <div id="container"></div>
  </div>
</template>

<script>
import parse from "@/core/services/htmlParser.service";
import doc from "@/core/services/doc.service";
// import { renderAsync } from "docx-preview";

// import domJSON from "domjson";
export default {
  name: "TestComponent",
  data() {
    return {
      file: null,
      text: null,
      document: null,
      parsedText: null,
    };
  },

  methods: {
    async submitForm() {
      this.document = await doc(this.parsedText);
      // renderAsync(this.document, document.getElementById("container")).then(
      //   () => console.log("docx: finished")
      // );
    },
    fileUploaded(event) {
      const [f] = event.target.files;
      this.file = f;
      const reader = new FileReader();
      reader.onload = (() => (e) => {
        this.text = e.target.result
          .replace(/\r?\n|\r/g, "")
          .replace(/\s\s+/g, " ")
          .replace(/>\s+</g, "><");
      })(f);
      // Read in the image file as a data URL.
      reader.readAsText(f);
    },
  },
  watch: {
    text(newVal) {
      if (newVal !== null) {
        this.parsedText = parse(
          newVal
            .replace(/\r?\n|\r/g, "")
            .replace(/\s\s+/g, " ")
            .replace(/>\s+</g, "><")
        );
        this.submitForm();
      }
    },
  },
};
</script>

<style lang="scss" scoped>
.input {
  padding: 20px;
  width: 100%;
  display: grid;
  place-items: center;
}
label {
  display: block;
  text-align: center;
  padding: 20px;
}
</style>
