Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    document.querySelectorAll(".shape-btn").forEach((btn) => {
      btn.addEventListener("click", () => insertShape(btn.dataset.shape));
    });
    document.querySelectorAll(".line-btn").forEach((btn) => {
      btn.addEventListener("click", () => insertLine());
    });
    document.getElementById("picture-input").addEventListener("change", handlePictureInsert);
  }
});

function showStatus(message) {
  document.getElementById("status").innerText = message;
}

// NOTE: PowerPoint.GeometricShapeType has more members than this. I've only
// used the ones I'm confident are correct spellings. Before relying on this,
// type "PowerPoint.GeometricShapeType." in VS Code with the Office.js types
// installed and check autocomplete — some of your original shapes (Diamond 2,
// Rhombus vs Diamond, Generic/Equilateral/Isosceles Triangle) may collapse to
// the same enum member, or may not exist as separate options at all.
async function insertShape(shapeKey) {
  try {
    await PowerPoint.run(async (context) => {
      const slides = context.presentation.getSelectedSlides();
      slides.load("items");
      await context.sync();

      if (slides.items.length === 0) {
        showStatus("Select a slide first.");
        return;
      }
      const slide = slides.items[0];

      let geometricType;
      let width = 150;
      let height = 150;

      switch (shapeKey) {
        case "rectangle":
          geometricType = PowerPoint.GeometricShapeType.rectangle;
          width = 200; height = 120;
          break;
        case "square":
          geometricType = PowerPoint.GeometricShapeType.rectangle;
          width = 150; height = 150;
          break;
        case "ellipse":
          geometricType = PowerPoint.GeometricShapeType.ellipse;
          width = 200; height = 120;
          break;
        case "circle":
          geometricType = PowerPoint.GeometricShapeType.ellipse;
          width = 150; height = 150;
          break;
        case "triangle":
          geometricType = PowerPoint.GeometricShapeType.triangle;
          break;
        case "rightTriangle":
          geometricType = PowerPoint.GeometricShapeType.rightTriangle;
          break;
        case "diamond":
          geometricType = PowerPoint.GeometricShapeType.diamond;
          break;
        case "parallelogram":
          geometricType = PowerPoint.GeometricShapeType.parallelogram;
          break;
        case "trapezoid":
          geometricType = PowerPoint.GeometricShapeType.trapezoid;
          break;
        case "hexagon":
          geometricType = PowerPoint.GeometricShapeType.hexagon;
          break;
        case "can":
          geometricType = PowerPoint.GeometricShapeType.can; // cylinder
          break;
        default:
          showStatus("Unknown shape: " + shapeKey);
          return;
      }

      const shape = slide.shapes.addGeometricShape(geometricType, {
        left: 100,
        top: 100,
        width: width,
        height: height,
      });

      shape.fill.clear();         // no fill
      shape.lineFormat.weight = 1; // 1pt outline
      shape.lineFormat.color = "#000000";

      // set default text formatting for anything typed into this shape
      shape.textFrame.textRange.font.name = "Comic Sans MS";
      shape.textFrame.textRange.font.color = "#000000";

      //horizontal centering
      shape.textFrame.textRange.paragraphFormat.horizontalAlignment = "Center";

      //vertical centering
      shape.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middleCentered;

      await context.sync();
      showStatus("Inserted " + shapeKey);
    });
  } catch (error) {
    showStatus("Error: " + error.message);
    console.error(error);
  }
}

async function insertLine() {
  try {
    await PowerPoint.run(async (context) => {
      const slides = context.presentation.getSelectedSlides();
      slides.load("items");
      await context.sync();

      if (slides.items.length === 0) {
        showStatus("Select a slide first.");
        return;
      }
      const slide = slides.items[0];

      const line = slide.shapes.addLine(PowerPoint.ConnectorType.straight, {
        left: 100,
        top: 150,
        width: 200,
        height: 0,
      });
      line.name = "Line";
      line.lineFormat.weight = 1;
      line.lineFormat.color = "#000000";

      await context.sync();
      showStatus("Inserted line");
    });
  } catch (error) {
    showStatus("Error: " + error.message);
    console.error(error);
  }
}

function handlePictureInsert(event) {
  const file = event.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = async function (e) {
    // result looks like "data:image/png;base64,AAAA..." — we need just the base64 part
    const base64 = e.target.result.split(",")[1];
    try {
      await PowerPoint.run(async (context) => {
        const slides = context.presentation.getSelectedSlides();
        slides.load("items");
        await context.sync();

        if (slides.items.length === 0) {
          showStatus("Select a slide first.");
          return;
        }
        const slide = slides.items[0];
        slide.shapes.addImage(base64, {
          left: 100,
          top: 100,
          width: 300,
          height: 200,
        });
        await context.sync();
        showStatus("Inserted picture: " + file.name);
      });
    } catch (error) {
      showStatus("Error: " + error.message);
      console.error(error);
    }
  };
  reader.readAsDataURL(file);
}
