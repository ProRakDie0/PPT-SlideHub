import slidesCatalog from "./slideData.json";

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    renderSlides(slidesCatalog);
    initSearch();
  }
});

function renderSlides(slides) {
  const container = document.getElementById("slides-container");
  if (!container) return;

  container.innerHTML = "";

  slides.forEach((slide) => {

    const previewSrc = slide.preview || "";

    const tagsHTML = slide.tags.map((tag) => `<span class="tag">${tag}</span>`).join("");

    const card = document.createElement("div");
    card.className = "card";
    card.innerHTML = `
      <img class="card-preview" src="${previewSrc}" alt="${slide.name}" />
      <div class="card-header">
        <h3 class="card-title">${slide.name}</h3>
        <button class="card-menu-btn">•••</button>
      </div>
      <div class="card-metadata">Размер: 16:9 • Слайдов: 1</div>
      <div class="card-tags">${tagsHTML}</div>
      <button class="insert-btn" data-url="${slide.fileUrl}">Insert Slide</button>
    `;

    const insertBtn = card.querySelector(".insert-btn");
    insertBtn.addEventListener("click", async () => {
      await handleInsertSlide(slide.fileUrl, insertBtn);
    });

    container.appendChild(card);
  });
}

async function handleInsertSlide(fileUrl, buttonEl) {
  const originalLabel = buttonEl ? buttonEl.textContent : null;

  try {
    if (buttonEl) {
      buttonEl.disabled = true;
      buttonEl.textContent = "Inserting...";
    }

    const response = await fetch(fileUrl);
    if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);

    const base64Data = (await response.text()).trim();

    await PowerPoint.run(async (context) => {
      context.presentation.insertSlidesFromBase64(base64Data, {
        formatting: "UseDestinationTheme",
      });
      await context.sync();
    });

    console.log("Слайд успешно вставлен в презентацию.");
  } catch (error) {
    console.error("Ошибка при загрузке или вставке слайда:", error);
  } finally {
    if (buttonEl) {
      buttonEl.disabled = false;
      buttonEl.textContent = originalLabel;
    }
  }
}

function initSearch() {
  const searchInput = document.getElementById("search-input");
  if (!searchInput) return;

  searchInput.addEventListener("input", (e) => {
    const query = e.target.value.toLowerCase().trim();

    const filtered = slidesCatalog.filter((slide) => {
      const matchName = slide.name.toLowerCase().includes(query);
      const matchTags = slide.tags.some((tag) => tag.toLowerCase().includes(query));
      return matchName || matchTags;
    });

    renderSlides(filtered);
  });
}