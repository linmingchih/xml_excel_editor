



    const state = {



      materials: [],



      materialMap: new Map(),



      materialIdCounter: 0,



      layers: [],



      lengthUnit: "",
      selectedMaterialName: ""



    };







    const dom = {



      fileInput: document.getElementById("fileInput"),



      downloadXml: document.getElementById("downloadXml"),



      materialsSection: document.getElementById("materialsSection"),



      materialsTableBody: document.querySelector("#materialsTable tbody"),



      layersSection: document.getElementById("layersSection"),



      layersTableBody: document.querySelector("#layersTable tbody"),






      lengthUnit: document.getElementById("lengthUnit"),



      otherLayersSection: document.getElementById("otherLayersSection"),



      otherLayersTableBody: document.querySelector("#otherLayersTable tbody")



    };







    const materialFieldDescriptors = [



      { key: "name" },



      { key: "permittivity" },



      { key: "lossTangent" },



      { key: "conductivity" }



    ];







    const layerFieldDescriptors = [



      { key: "name" },



      { key: "type" },



      { key: "material" },



      { key: "materialPermittivity", materialTarget: "material", materialProp: "permittivity" },



      { key: "materialLossTangent", materialTarget: "material", materialProp: "lossTangent" },



      { key: "materialConductivity", materialTarget: "material", materialProp: "conductivity" },



      { key: "fillMaterial" },



      { key: "thickness" },



      { key: "etchFactor" },



      { key: "color" },



      { key: "hurayTopRatio" },



      { key: "hurayTopRadius" },



      { key: "hurayBottomRatio" },



      { key: "hurayBottomRadius" },



      { key: "huraySideRatio" },



      { key: "huraySideRadius" }



    ];







    const layerFieldIndex = new Map(layerFieldDescriptors.map((descriptor, idx) => [descriptor.key, idx]));







    dom.fileInput.addEventListener("change", async (event) => {



      const files = event.target.files;



      const file = files && files[0];



      if (!file) {



        return;



      }



      try {



        const text = await file.text();



        loadXml(text);



      } catch (error) {



        console.error(error);



        alert("Failed to read file. Please check browser permissions.");



      }



    });







    dom.downloadXml.addEventListener("click", () => {



      try {



        const xmlString = buildXml();



        const blob = new Blob([xmlString], { type: "application/xml" });



        const link = document.createElement("a");



        link.href = URL.createObjectURL(blob);



        link.download = "stackup_export.xml";



        document.body.appendChild(link);



        link.click();



        document.body.removeChild(link);



        URL.revokeObjectURL(link.href);



      } catch (error) {



        console.error(error);



        alert("Error while generating XML. Please review the table for missing values.");



      }



    });















    dom.lengthUnit.addEventListener("input", (event) => {



      state.lengthUnit = event.target.value.trim();



    });
    dom.materialsTableBody.addEventListener("click", (event) => {
      const target = event.target;
      if (!(target instanceof HTMLElement)) {
        return;
      }
      const row = target.closest("tr");
      if (!row) {
        return;
      }
      const index = Number(row.dataset.index || "-1");
      if (Number.isNaN(index) || index < 0) {
        return;
      }
      const material = state.materials[index];
      if (!material) {
        return;
      }
      const normalizedName = normalizeMaterialPropValue(material.name);
      if (state.selectedMaterialName && normalizedName === normalizeMaterialPropValue(state.selectedMaterialName)) {
        state.selectedMaterialName = "";
      } else {
        state.selectedMaterialName = normalizedName;
      }
      renderMaterials();
      renderLayers();
    });








    document.body.addEventListener("input", (event) => {



      const element = event.target;



      if (!(element instanceof HTMLElement)) {



        return;



      }



      const collection = element.dataset.collection;



      if (!collection) {



        return;



      }







      if (!(element instanceof HTMLInputElement)) {



        sanitizeCell(element);



      }







      const rowElement = element.closest("tr");



      if (!rowElement) {



        return;



      }



      const index = Number(rowElement.dataset.index || "-1");



      if (Number.isNaN(index) || index < 0) {



        return;



      }



      const field = element.dataset.field;



      if (!field) {



        return;



      }







      const value = element instanceof HTMLInputElement ? element.value : element.textContent.trim();







      if (collection === "materials") {



        updateMaterialCell(index, field, value);



        renderLayers();



      } else if (collection === "layers") {



        updateLayerCell(index, field, value, { ensureMaterial: false });



        const layer = state.layers[index];



        if (!layer) {



          return;



        }



        if (field === "etchFactor" && !(element instanceof HTMLInputElement)) {



          element.textContent = layer.etchFactor || "";



        }



        if (field === "material" && !(element instanceof HTMLInputElement)) {



          element.textContent = layer.material || "";



        }



        if (field === "fillMaterial" && !(element instanceof HTMLInputElement)) {



          element.textContent = layer.fillMaterial || "";



        }



        if (field === "color" && element instanceof HTMLInputElement) {



          element.value = ensureColorInputValue(layer.color);



          element.title = ensureColorInputValue(layer.color);



        }



      }



    });







    document.body.addEventListener("focusout", (event) => {



      const cell = event.target;



      if (!(cell instanceof HTMLElement)) {



        return;



      }



      if (cell.dataset.collection !== "layers") {



        return;



      }



      const rowElement = cell.closest("tr");



      if (!rowElement) {



        return;



      }



      const index = Number(rowElement.dataset.index || "-1");



      if (Number.isNaN(index) || index < 0) {



        return;



      }



      const layer = state.layers[index];



      if (!layer) {



        return;



      }



      const field = cell.dataset.field;



      if (field === "material") {



        commitLayerMaterial(layer, "material");



        renderLayers();



      } else if (field === "fillMaterial") {



        commitLayerMaterial(layer, "fillMaterial");



        renderLayers();



      } else if (field === "type") {



        renderLayers();



      } else if (cell.dataset.materialProp) {



        renderLayers();



      }



    }, true);







    document.body.addEventListener("keydown", (event) => {



      const cell = event.target;



      if (!(cell instanceof HTMLElement)) {



        return;



      }



      if (cell.dataset.collection !== "layers") {



        return;



      }



      if (event.key === "Enter") {



        event.preventDefault();



        moveFocus(cell, event.shiftKey ? -1 : 1);



      }



    });







    document.body.addEventListener("paste", (event) => {



      const cell = event.target;



      if (!(cell instanceof HTMLElement)) {



        return;



      }



      const collection = cell.dataset.collection;



      if (!collection) {



        return;



      }



      const clipboard = event.clipboardData && event.clipboardData.getData("text");



      if (!clipboard || (clipboard.indexOf("\t") === -1 && clipboard.indexOf("\n") === -1)) {



        return;



      }



      event.preventDefault();



      const matrix = clipboard



        .split(/\r?\n/)



        .filter((row, idx, arr) => !(row === "" && idx === arr.length - 1))



        .map((row) => row.split("\t"));



      const rowElement = cell.closest("tr");



      if (!rowElement) {



        return;



      }



      const startRow = Number(rowElement.dataset.index || "-1");



      if (Number.isNaN(startRow) || startRow < 0) {



        return;



      }



      const field = cell.dataset.field;



      if (!field) {



        return;



      }



      if (collection === "layers") {



        const startField = layerFieldIndex.get(field);



        if (startField === undefined) {



          return;



        }



        applyMatrixToLayers(startRow, startField, matrix);



      } else if (collection === "materials") {



        const startField = materialFieldDescriptors.findIndex((descriptor) => descriptor.key === field);



        if (startField === -1) {



          return;



        }



        applyMatrixToMaterials(startRow, startField, matrix);



      }



    });







    document.body.addEventListener("click", (event) => {



      const button = event.target;



      if (!(button instanceof HTMLButtonElement)) {



        return;



      }



      const action = button.dataset.action;



      if (!action) {



        return;



      }



      const rowElement = button.closest("tr");



      if (!rowElement) {



        return;



      }



      const index = Number(rowElement.dataset.index || "-1");



      if (Number.isNaN(index) || index < 0) {



        return;



      }



      if (action === "remove-layer") {



        state.layers.splice(index, 1);



        renderLayers();



      } else if (action === "remove-material") {



        removeMaterialAtIndex(index);



        renderLayers();



      }



    });



    function loadXml(xmlText) {



      resetState();



      const parser = new DOMParser();



      const doc = parser.parseFromString(xmlText, "application/xml");



      const errorNode = doc.getElementsByTagName("parsererror")[0];



      if (errorNode) {



        console.error(errorNode.textContent || "parsererror");



        alert("Unable to parse XML. Please check the file format.");



        return;



      }







      const control = doc.documentElement;



      if (!control || control.localName !== "Control") {



        alert("Expected c:Control root element.");



        return;



      }







      const stackup = findChildElement(control, "Stackup");



      if (!stackup) {



        alert("Unable to locate Stackup element in the XML.");



        return;



      }







      parseMaterials(stackup);



      parseLayers(stackup);







      dom.materialsSection.classList.remove("hidden");



      dom.layersSection.classList.remove("hidden");



      dom.downloadXml.disabled = false;



      dom.lengthUnit.value = state.lengthUnit;







      pruneUnusedMaterials();
      renderMaterials();



      renderLayers();



    }







    function resetState() {



      state.materials = [];



      state.materialMap = new Map();



      state.materialIdCounter = 0;



      state.layers = [];



      state.lengthUnit = "";

      state.selectedMaterialName = "";


    }







    function parseMaterials(stackup) {



      const materialsParent = findChildElement(stackup, "Materials");



      const materialNodes = findChildElements(materialsParent, "Material");



      materialNodes.forEach((node) => {



        const name = (node.getAttribute("Name") || "").trim();



        const entry = {



          id: generateMaterialId(),



          name: name,



          permittivity: readDouble(node, "Permittivity"),



          lossTangent: readDouble(node, "DielectricLossTangent"),



          conductivity: readDouble(node, "Conductivity"),
          aliases: []
        };



        state.materials.push(entry);



        if (name) {
          registerMaterialAlias(entry, name);
        }



      });



    }







    function parseLayers(stackup) {



      const layersRoot = findChildElement(stackup, "Layers");



      if (!layersRoot) {



        return;



      }



      const lengthUnitAttr = layersRoot.getAttribute("LengthUnit");



      state.lengthUnit = lengthUnitAttr ? lengthUnitAttr : "";



      const layerNodes = findChildElements(layersRoot, "Layer");



      layerNodes.forEach((node) => {



        const layer = createEmptyLayer();



        layer.name = node.getAttribute("Name") || "";



        layer.type = node.getAttribute("Type") || "";



        layer.material = node.getAttribute("Material") || "";



        layer.fillMaterial = node.getAttribute("FillMaterial") || "";



        layer.thickness = node.getAttribute("Thickness") || "";



        layer.etchFactor = formatEtchFactor(node.getAttribute("EtchFactor") || "");



        layer.color = normalizeColorValue(node.getAttribute("Color") || "") || "";



        layer.hurayTopRatio = readHuray(node, "HuraySurfaceRoughness", "HallHuraySurfaceRatio");



        layer.hurayTopRadius = readHuray(node, "HuraySurfaceRoughness", "NoduleRadius");



        layer.hurayBottomRatio = readHuray(node, "HurayBottomSurfaceRoughness", "HallHuraySurfaceRatio");



        layer.hurayBottomRadius = readHuray(node, "HurayBottomSurfaceRoughness", "NoduleRadius");



        layer.huraySideRatio = readHuray(node, "HuraySideSurfaceRoughness", "HallHuraySurfaceRatio");



        layer.huraySideRadius = readHuray(node, "HuraySideSurfaceRoughness", "NoduleRadius");







        assignProps(layer.materialPropsFallback, findMaterial(layer.material));



        assignProps(layer.fillMaterialPropsFallback, findMaterial(layer.fillMaterial));







        state.layers.push(layer);



      });



    }







    function createEmptyLayer() {



      return {



        name: "",



        type: "",



        material: "",



        fillMaterial: "",



        thickness: "",



        etchFactor: "",



        color: "",



        hurayTopRatio: "",



        hurayTopRadius: "",



        hurayBottomRatio: "",



        hurayBottomRadius: "",



        huraySideRatio: "",



        huraySideRadius: "",



        materialPropsFallback: { permittivity: "", lossTangent: "", conductivity: "" },



        fillMaterialPropsFallback: { permittivity: "", lossTangent: "", conductivity: "" }



      };



    }







    function addMaterial(initial = {}) {
      const desiredName = typeof initial.name === "string" ? initial.name.trim() : "";
      const entry = {
        id: generateMaterialId(),
        name: ensureUniqueMaterialName(desiredName),
        permittivity: initial.permittivity ?? "",
        lossTangent: initial.lossTangent ?? "",
        conductivity: initial.conductivity ?? "",
        aliases: []
      };
      state.materials.push(entry);
      registerMaterialAlias(entry, entry.name);
      return entry;
    }

    function registerMaterialAlias(entry, alias) {
      const normalized = normalizeMaterialPropValue(alias);
      if (!normalized) {
        return;
      }
      if (!entry.aliases) {
        entry.aliases = [];
      }
      if (!entry.aliases.includes(normalized)) {
        entry.aliases.push(normalized);
      }
      state.materialMap.set(normalized, entry);
      state.materialMap.set(normalized.toLowerCase(), entry);
      state.materialMap.set(normalized.toUpperCase(), entry);
    }

    function unregisterMaterialAlias(name) {
      const normalized = normalizeMaterialPropValue(name);
      if (!normalized) {
        return;
      }
      state.materialMap.delete(normalized);
      state.materialMap.delete(normalized.toLowerCase());
      state.materialMap.delete(normalized.toUpperCase());
    }

    function unregisterMaterialAliases(entry) {
      if (!entry || !entry.aliases) {
        return;
      }
      entry.aliases.forEach((alias) => {
        unregisterMaterialAlias(alias);
      });
      entry.aliases.length = 0;
    }

    function transferMaterialAliases(source, target) {
      if (!source || !source.aliases || !target) {
        return;
      }
      source.aliases.forEach((alias) => {
        registerMaterialAlias(target, alias);
      });
      source.aliases.length = 0;
    }

    function ensureUniqueMaterialName(baseName) {
      const trimmed = (baseName || "").trim();
      if (!trimmed) {
        return generateMaterialName();
      }
      let candidate = trimmed;
      let suffix = 1;
      while (isMaterialNameTaken(candidate)) {
        candidate = `${trimmed}_${suffix}`;
        suffix += 1;
      }
      return candidate;
    }

    function isMaterialNameTaken(name) {
      if (!name) {
        return false;
      }
      return state.materialMap.has(name)
        || state.materialMap.has(name.toLowerCase())
        || state.materialMap.has(name.toUpperCase());
    }







    function renderMaterials() {
      const selected = normalizeMaterialPropValue(state.selectedMaterialName);
      dom.materialsTableBody.innerHTML = "";
      state.materials.forEach((material, index) => {
        const row = document.createElement("tr");
        row.dataset.index = String(index);
        row.dataset.name = material.name || "";
        row.classList.add("material-row");
        const materialName = normalizeMaterialPropValue(material.name);
        if (selected && materialName && materialName === selected) {
          row.classList.add("selected");
        }
        materialFieldDescriptors.forEach((descriptor) => {
          const cell = document.createElement("td");
          const value = material[descriptor.key] || "";
          cell.textContent = value;
          if (descriptor.key !== "name") {
            cell.classList.add("numeric");
          }
          row.appendChild(cell);
        });
        dom.materialsTableBody.appendChild(row);
      });
      if (state.materials.length === 0) {
        state.selectedMaterialName = "";
      }
    }

    function pruneUnusedMaterials() {
      let modified = false;
      for (let index = state.materials.length - 1; index >= 0; index -= 1) {
        const material = state.materials[index];
        if (!material) {
          continue;
        }
        const materialName = material.name || "";
        if (countMaterialAssignments(materialName) === 0) {
          unregisterMaterialAliases(material);
          state.materials.splice(index, 1);
          modified = true;
        }
      }
      if (modified) {
        const normalizedSelected = normalizeMaterialPropValue(state.selectedMaterialName);
        if (normalizedSelected && !findMaterial(normalizedSelected)) {
          state.selectedMaterialName = "";
        }
      }
      return modified;
    }

    function renderLayers() {



      autoAssignFillMaterials();



      dom.layersTableBody.innerHTML = "";
      const selectedMaterial = normalizeMaterialPropValue(state.selectedMaterialName);



      let hasEditable = false;







      state.layers.forEach((layer, index) => {



        if (!isConductor(layer.type) && !isDielectric(layer.type)) {



          return;



        }



        hasEditable = true;



        const row = document.createElement("tr");
        row.classList.add("layer-row");
        const rowMaterialName = normalizeMaterialPropValue(layer.material);
        const rowFillMaterialName = normalizeMaterialPropValue(layer.fillMaterial);
        if (isConductor(layer.type)) {
          row.classList.add("conductor-row");
        } else if (isDielectric(layer.type)) {
          row.classList.add("dielectric-row");
        }
        row.dataset.index = String(index);
        row.dataset.materialName = layer.material || "";
        row.dataset.fillMaterialName = layer.fillMaterial || "";
        if (selectedMaterial && (rowMaterialName === selectedMaterial || rowFillMaterialName === selectedMaterial)) {
          row.classList.add("layer-highlight");
        }







        layerFieldDescriptors.forEach((descriptor) => {



          const cell = document.createElement("td");







          if (descriptor.materialProp) {



            cell.contentEditable = "true";



            cell.tabIndex = 0;



            cell.dataset.collection = "layers";



            cell.dataset.field = descriptor.key;



            cell.dataset.materialTarget = descriptor.materialTarget;



            cell.dataset.materialProp = descriptor.materialProp;



            cell.spellcheck = false;



            cell.classList.add("numeric");



            cell.textContent = getLayerMaterialValue(layer, descriptor.materialTarget, descriptor.materialProp);



          } else if (descriptor.key === "color") {



            cell.classList.add("color");



            const colorInput = document.createElement("input");



            colorInput.type = "color";



            colorInput.dataset.collection = "layers";



            colorInput.dataset.field = descriptor.key;



            const colorValue = ensureColorInputValue(layer.color);



            colorInput.value = colorValue;



            colorInput.title = layer.color || colorValue;



            cell.appendChild(colorInput);



          } else {



            cell.contentEditable = "true";



            cell.tabIndex = 0;



            cell.dataset.collection = "layers";



            cell.dataset.field = descriptor.key;



            cell.spellcheck = false;



            cell.textContent = descriptor.key === "etchFactor" ? layer.etchFactor || "" : layer[descriptor.key] || "";



          }







          row.appendChild(cell);



        });







        dom.layersTableBody.appendChild(row);



      });







      if (hasEditable) {



        dom.layersSection.classList.remove("hidden");



      } else {



        dom.layersSection.classList.add("hidden");



      }







      renderOtherLayers();
      pruneUnusedMaterials();
      renderMaterials();



    }







    function renderOtherLayers() {
      const selectedMaterial = normalizeMaterialPropValue(state.selectedMaterialName);



      const tbody = dom.otherLayersTableBody;



      if (!tbody) {



        return;



      }



      tbody.innerHTML = "";



      const otherLayers = state.layers



        .map((layer, layerIndex) => ({ layer, layerIndex }))



        .filter(({ layer }) => !isConductor(layer.type) && !isDielectric(layer.type));



      if (otherLayers.length === 0) {



        dom.otherLayersSection.classList.add("hidden");



        return;



      }



      dom.otherLayersSection.classList.remove("hidden");



      otherLayers.forEach(({ layer, layerIndex }) => {



        const row = document.createElement("tr");
        row.classList.add("other-layer-row");
        const normalizedMaterial = normalizeMaterialPropValue(layer.material);
        const normalizedFill = normalizeMaterialPropValue(layer.fillMaterial);
        if (selectedMaterial && (normalizedMaterial === selectedMaterial || normalizedFill === selectedMaterial)) {
          row.classList.add("layer-highlight");
        }
        row.dataset.index = String(layerIndex);
        row.dataset.materialName = layer.material || "";
        row.dataset.fillMaterialName = layer.fillMaterial || "";







        const nameCell = document.createElement("td");



        nameCell.textContent = layer.name || "";



        row.appendChild(nameCell);







        const typeCell = document.createElement("td");



        typeCell.textContent = layer.type || "";



        row.appendChild(typeCell);







        const materialCell = document.createElement("td");



        materialCell.textContent = layer.material || "";



        row.appendChild(materialCell);







        const thicknessCell = document.createElement("td");



        thicknessCell.classList.add("numeric");



        thicknessCell.textContent = layer.thickness || "";



        row.appendChild(thicknessCell);







        const colorCell = document.createElement("td");



        colorCell.classList.add("color");



        const normalizedColor = normalizeColorValue(layer.color || "");



        if (normalizedColor) {



          const swatch = document.createElement("span");



          swatch.style.display = "inline-block";



          swatch.style.width = "1.5rem";



          swatch.style.height = "1rem";



          swatch.style.marginRight = "0.5rem";



          swatch.style.border = "1px solid #c7d2e2";



          swatch.style.background = normalizedColor;



          swatch.title = normalizedColor;



          colorCell.appendChild(swatch);



          const label = document.createElement("span");



          label.textContent = normalizedColor;



          colorCell.appendChild(label);



        } else {



          colorCell.textContent = "";



        }



        row.appendChild(colorCell);







        tbody.appendChild(row);



      });



    }







    function updateMaterialCell(index, field, value) {}







    function updateLayerCell(index, field, value, options) {



      const layer = createLayerAt(index);



      switch (field) {



        case "name":



        case "type":



        case "thickness":



        case "hurayTopRatio":



        case "hurayTopRadius":



        case "hurayBottomRatio":



        case "hurayBottomRadius":



        case "huraySideRatio":



        case "huraySideRadius":



          layer[field] = value;



          break;



        case "etchFactor":



          layer.etchFactor = formatEtchFactor(value);



          break;



        case "color":



          layer.color = normalizeColorValue(value) || ensureColorInputValue(value);



          break;



        case "material": {



          const entry = ensureMaterial(value);



          if (entry) {



            layer.material = entry.name;



            assignProps(layer.materialPropsFallback, entry);



          } else {



            layer.material = "";



            assignProps(layer.materialPropsFallback, null);



          }



          break;



        }



        case "fillMaterial": {



          const entry = ensureMaterial(value);



          if (entry) {



            layer.fillMaterial = entry.name;



            assignProps(layer.fillMaterialPropsFallback, entry);



          } else {



            layer.fillMaterial = "";



            assignProps(layer.fillMaterialPropsFallback, null);



          }



          break;



        }



        case "materialPermittivity":



          setLayerMaterialProperty(layer, "material", "permittivity", value, true);



          break;



        case "materialLossTangent":



          setLayerMaterialProperty(layer, "material", "lossTangent", value, true);



          break;



        case "materialConductivity":



          setLayerMaterialProperty(layer, "material", "conductivity", value, true);



          break;



        default:



          break;



      }



    }







    function createLayerAt(index) {



      while (state.layers.length <= index) {



        state.layers.push(createEmptyLayer());



      }



      return state.layers[index];



    }







    function renameMaterial(index, newName) {
      const material = state.materials[index];
      if (!material) {
        return;
      }
      const trimmed = normalizeMaterialPropValue(newName);
      const oldName = material.name;
      if (!trimmed) {
        unregisterMaterialAliases(material);
        material.name = "";
        propagateMaterialDeletion(oldName);
        return;
      }
      if (trimmed === oldName) {
        registerMaterialAlias(material, trimmed);
        return;
      }
      const existing = findMaterial(trimmed);
      if (existing && existing !== material) {
        mergeMaterialEntries(material, existing);
        transferMaterialAliases(material, existing);
        state.materials.splice(index, 1);
        updateLayersMaterialName(oldName, existing.name);
      } else {
        unregisterMaterialAliases(material);
        material.name = trimmed;
        registerMaterialAlias(material, trimmed);
        updateLayersMaterialName(oldName, trimmed);
      }
    }

    function mergeMaterialEntries(source, target) {



      ["permittivity", "lossTangent", "conductivity"].forEach((prop) => {



        if (!target[prop] && source[prop]) {



          target[prop] = source[prop];



        }



      });



    }







    function updateLayersMaterialName(oldName, newName) {



      state.layers.forEach((layer) => {



        if (oldName && layer.material === oldName) {



          layer.material = newName;



        }



        if (oldName && layer.fillMaterial === oldName) {



          layer.fillMaterial = newName;



        }



      });



    }







    function propagateMaterialToLayers(material) {



      state.layers.forEach((layer) => {



        if (material.name && layer.material === material.name) {



          assignProps(layer.materialPropsFallback, material);



        }



        if (material.name && layer.fillMaterial === material.name) {



          assignProps(layer.fillMaterialPropsFallback, material);



        }



      });



    }







    function propagateMaterialDeletion(oldName) {



      state.layers.forEach((layer) => {



        if (layer.material === oldName) {



          layer.material = "";



        }



        if (layer.fillMaterial === oldName) {



          layer.fillMaterial = "";



        }



      });



    }







    function removeMaterialAtIndex(index) {
      const material = state.materials[index];
      if (!material) {
        return;
      }
      const oldName = material.name;
      unregisterMaterialAliases(material);
      state.materials.splice(index, 1);
      propagateMaterialDeletion(oldName);
    }

    function sanitizeCell(cell) {



      const cleaned = cell.textContent.replace(/\u00a0/g, " ").replace(/\r?\n+/g, " ");



      if (cleaned !== cell.textContent) {



        cell.textContent = cleaned;



        placeCaretAtEnd(cell);



      }



    }







    function setLayerMaterialProperty(layer, target, prop, value, ensure) {
      const fallbackKey = target === "material" ? "materialPropsFallback" : "fillMaterialPropsFallback";
      const normalizedValue = normalizeMaterialPropValue(value);
      layer[fallbackKey][prop] = normalizedValue;

      if (!ensure || (target !== "material" && target !== "fillMaterial")) {
        return;
      }

      const desiredProps = {
        permittivity: normalizeMaterialPropValue(layer[fallbackKey].permittivity),
        lossTangent: normalizeMaterialPropValue(layer[fallbackKey].lossTangent),
        conductivity: normalizeMaterialPropValue(layer[fallbackKey].conductivity)
      };

      const hasAnyValue = desiredProps.permittivity || desiredProps.lossTangent || desiredProps.conductivity;

      if (!hasAnyValue) {
        assignLayerMaterial(layer, target, null);
        return;
      }

      const currentName = target === "material" ? layer.material : layer.fillMaterial;
      const currentEntry = currentName ? findMaterial(currentName) : null;
      const existing = findMaterialByProperties(desiredProps);
      if (existing) {
        assignLayerMaterial(layer, target, existing);
        renderMaterials();
        return;
      }

      if (currentEntry && countMaterialAssignments(currentEntry.name) === 1) {
        const oldName = currentEntry.name;
        currentEntry.permittivity = desiredProps.permittivity;
        currentEntry.lossTangent = desiredProps.lossTangent;
        currentEntry.conductivity = desiredProps.conductivity;
        const generatedName = ensureUniqueMaterialName(buildMaterialNameFromProps(desiredProps));
        if (generatedName !== oldName) {
          unregisterMaterialAliases(currentEntry);
          currentEntry.name = generatedName;
          registerMaterialAlias(currentEntry, generatedName);
          updateLayersMaterialName(oldName, generatedName);
        } else {
          registerMaterialAlias(currentEntry, generatedName);
        }
        assignLayerMaterial(layer, target, currentEntry);
        renderMaterials();
        return;
      }

      const entry = addMaterial({
        name: buildMaterialNameFromProps(desiredProps),
        permittivity: desiredProps.permittivity,
        lossTangent: desiredProps.lossTangent,
        conductivity: desiredProps.conductivity
      });

      assignLayerMaterial(layer, target, entry);
    }

    function ensureMaterial(name) {
      const trimmed = normalizeMaterialPropValue(name);
      if (trimmed) {
        const existing = findMaterial(trimmed);
        if (existing) {
          return existing;
        }
        const created = addMaterial({ name: trimmed });
        registerMaterialAlias(created, trimmed);
        return created;
      }
      return addMaterial();
    }

    function findMaterial(name) {
      const trimmed = normalizeMaterialPropValue(name);
      if (!trimmed) {
        return null;
      }
      const lower = trimmed.toLowerCase();
      const upper = trimmed.toUpperCase();
      return state.materialMap.get(trimmed)
        || state.materialMap.get(lower)
        || state.materialMap.get(upper)
        || null;
    }

    function normalizeMaterialPropValue(value) {
      if (value === undefined || value === null) {
        return "";
      }
      return String(value).trim();
    }

    function findMaterialByProperties(props) {
      for (let index = 0; index < state.materials.length; index += 1) {
        const material = state.materials[index];
        if (normalizeMaterialPropValue(material.permittivity) === props.permittivity
          && normalizeMaterialPropValue(material.lossTangent) === props.lossTangent
          && normalizeMaterialPropValue(material.conductivity) === props.conductivity) {
          return material;
        }
      }
      return null;
    }

    function sanitizeMaterialNamePart(value) {
      const trimmed = normalizeMaterialPropValue(value);
      if (!trimmed) {
        return "x";
      }
      const cleaned = trimmed.replace(/[^0-9a-zA-Z.\-]/g, "");
      return cleaned || "x";
    }

    function buildMaterialNameFromProps(props) {
      const parts = [];
      if (props.permittivity) {
        parts.push(sanitizeMaterialNamePart(props.permittivity));
      }
      if (props.lossTangent) {
        parts.push(sanitizeMaterialNamePart(props.lossTangent));
      }
      if (props.conductivity) {
        parts.push(sanitizeMaterialNamePart(props.conductivity));
      }
      if (parts.length === 0) {
        return "m";
      }
      return `m_${parts.join("_")}`;
    }
    function countMaterialAssignments(name) {
      const normalizedName = normalizeMaterialPropValue(name);
      if (!normalizedName) {
        return 0;
      }
      let count = 0;
      for (let index = 0; index < state.layers.length; index += 1) {
        const candidate = state.layers[index];
        if (normalizeMaterialPropValue(candidate.material) === normalizedName) {
          count += 1;
        }
        if (normalizeMaterialPropValue(candidate.fillMaterial) === normalizedName) {
          count += 1;
        }
      }
      return count;
    }



    function assignLayerMaterial(layer, target, entry) {
      const fallbackKey = target === "material" ? "materialPropsFallback" : "fillMaterialPropsFallback";
      if (!entry) {
        if (target === "material") {
          layer.material = "";
        } else {
          layer.fillMaterial = "";
        }
        assignProps(layer[fallbackKey], null);
        return;
      }
      if (target === "material") {
        layer.material = entry.name;
      } else {
        layer.fillMaterial = entry.name;
      }
      assignProps(layer[fallbackKey], entry);
    }

    function commitLayerMaterial(layer, target) {
      const name = target === "material" ? layer.material : layer.fillMaterial;
      const trimmed = normalizeMaterialPropValue(name);
      if (!trimmed) {
        assignLayerMaterial(layer, target, null);
        return;
      }
      const entry = ensureMaterial(trimmed);
      assignLayerMaterial(layer, target, entry);
    }

    function getLayerMaterialValue(layer, target, prop) {



      const name = target === "material" ? layer.material : layer.fillMaterial;



      const entry = findMaterial(name);



      if (entry && entry[prop]) {



        return entry[prop];



      }



      const fallback = target === "material" ? layer.materialPropsFallback : layer.fillMaterialPropsFallback;



      return fallback[prop] || "";



    }







    function assignProps(target, source) {



      if (!target) {



        return;



      }



      if (!source) {



        target.permittivity = "";



        target.lossTangent = "";



        target.conductivity = "";



        return;



      }



      target.permittivity = source.permittivity || "";



      target.lossTangent = source.lossTangent || "";



      target.conductivity = source.conductivity || "";



    }



    function formatEtchFactor(value) {



      const trimmed = (value || "").trim();



      if (!trimmed) {



        return "";



      }



      const number = Number(trimmed);



      if (!Number.isFinite(number)) {



        return trimmed;



      }



      return number.toFixed(3);



    }







    function normalizeColorValue(value) {



      const trimmed = (value || "").trim();



      if (/^#[0-9a-fA-F]{6}$/.test(trimmed)) {



        return trimmed.toLowerCase();



      }



      if (/^[0-9a-fA-F]{6}$/.test(trimmed)) {



        return `#${trimmed.toLowerCase()}`;



      }



      return "";



    }







    function ensureColorInputValue(value) {



      return normalizeColorValue(value) || "#000000";



    }







    function autoAssignFillMaterials() {



      for (let index = 0; index < state.layers.length; index += 1) {



        const layer = state.layers[index];



        if (!isConductor(layer.type)) {



          continue;



        }



        const dielectric = findDielectricAbove(index);



        if (!dielectric) {



          continue;



        }



        const entry = ensureMaterial(dielectric.material);



        if (entry) {



          if (layer.fillMaterial !== entry.name) {



            layer.fillMaterial = entry.name;



          }



          assignProps(layer.fillMaterialPropsFallback, entry);



        } else {



          layer.fillMaterial = "";



          assignProps(layer.fillMaterialPropsFallback, null);



        }



      }



    }



    function findDielectricAbove(index) {



      for (let i = index - 1; i >= 0; i -= 1) {



        const candidate = state.layers[i];



        if (isDielectric(candidate.type)) {



          return candidate;



        }



      }



      return null;



    }







    function isConductor(type) {



      return normalizeLayerType(type) === "conductor";



    }







    function isDielectric(type) {



      return normalizeLayerType(type) === "dielectric";



    }







    function normalizeLayerType(type) {



      return (type || "").trim().toLowerCase();



    }







    function generateMaterialName() {



      const prefix = "MAT";



      let counter = 1;



      let candidate = `${prefix}${String(counter).padStart(3, "0")}`;



      while (state.materialMap.has(candidate)) {



        counter += 1;



        candidate = `${prefix}${String(counter).padStart(3, "0")}`;



      }



      return candidate;



    }







    function generateMaterialId() {







      state.materialIdCounter += 1;



      return "mat_" + state.materialIdCounter;



    }







    function placeCaretAtEnd(element) {



      const range = document.createRange();



      range.selectNodeContents(element);



      range.collapse(false);



      const selection = window.getSelection();



      if (!selection) {



        return;



      }



      selection.removeAllRanges();



      selection.addRange(range);



    }







    function moveFocus(currentCell, rowOffset) {



      const rowElement = currentCell.closest("tr");



      if (!rowElement) {



        return;



      }



      const index = Number(rowElement.dataset.index || "-1");



      if (Number.isNaN(index) || index < 0) {



        return;



      }



      const field = currentCell.dataset.field;



      if (!field) {



        return;



      }



      const targetRow = index + rowOffset;



      if (targetRow < 0) {



        return;



      }



      if (targetRow >= state.layers.length) {



        while (state.layers.length <= targetRow) {



          state.layers.push(createEmptyLayer());



        }



        renderLayers();



      }



      const selector = 'tr[data-index="' + targetRow + '"] td[data-field="' + field + '"]';



      const targetCell = dom.layersTableBody.querySelector(selector);



      if (targetCell) {



        targetCell.focus();



      }



    }



    function applyMatrixToLayers(startRow, startField, matrix) {



      for (let r = 0; r < matrix.length; r += 1) {



        const rowIndex = startRow + r;



        const row = createLayerAt(rowIndex);



        const rowData = matrix[r];



        for (let c = 0; c < rowData.length; c += 1) {



          const fieldIndex = startField + c;



          if (fieldIndex >= layerFieldDescriptors.length) {



            break;



          }



          const descriptor = layerFieldDescriptors[fieldIndex];



          const value = rowData[c] ? rowData[c].trim() : "";



          updateLayerCell(rowIndex, descriptor.key, value, { ensureMaterial: true });



        }



        commitLayerMaterial(row, "material");



        commitLayerMaterial(row, "fillMaterial");



      }



      renderMaterials();



      renderLayers();



    }







    function applyMatrixToMaterials(startRow, startField, matrix) {



      for (let r = 0; r < matrix.length; r += 1) {



        const rowIndex = startRow + r;



        while (state.materials.length <= rowIndex) {



          addMaterial();



        }



        const rowData = matrix[r];



        for (let c = 0; c < rowData.length; c += 1) {



          const fieldIndex = startField + c;



          if (fieldIndex >= materialFieldDescriptors.length) {



            break;



          }



          const descriptor = materialFieldDescriptors[fieldIndex];



          const value = rowData[c] ? rowData[c].trim() : "";



          updateMaterialCell(rowIndex, descriptor.key, value);



        }



      }



      renderMaterials();



      renderLayers();



    }







    function readDouble(parent, childTag) {



      const container = findChildElement(parent, childTag);



      if (!container) {



        return "";



      }



      const doubleNode = findChildElement(container, "Double");



      return doubleNode && doubleNode.textContent ? doubleNode.textContent.trim() : "";



    }







    function readHuray(layerNode, tagName, attribute) {



      const node = findChildElement(layerNode, tagName);



      if (!node) {



        return "";



      }



      const attr = node.getAttribute(attribute);



      return attr ? attr : "";



    }







    function buildXml() {



      const namespace = "http://www.ansys.com/control";



      const doc = document.implementation.createDocument(namespace, "c:Control", null);



      const control = doc.documentElement;



      control.setAttribute("schemaVersion", "1.0");







      const stackup = doc.createElement("Stackup");



      stackup.setAttribute("schemaVersion", "1.0");



      control.appendChild(stackup);







      const materials = doc.createElement("Materials");



      stackup.appendChild(materials);







      const uniqueMaterials = new Map();



      state.materials.forEach((material) => {



        const name = (material.name || "").trim();



        if (!name || uniqueMaterials.has(name)) {



          return;



        }



        uniqueMaterials.set(name, true);



        const materialNode = doc.createElement("Material");



        materialNode.setAttribute("Name", name);



        appendDoubleElement(doc, materialNode, "Permittivity", material.permittivity);



        appendDoubleElement(doc, materialNode, "DielectricLossTangent", material.lossTangent);



        appendDoubleElement(doc, materialNode, "Conductivity", material.conductivity);



        materials.appendChild(materialNode);



      });







      state.layers.forEach((layer) => {



        ["material", "fillMaterial"].forEach((target) => {



          const name = (layer[target] || "").trim();



          if (!name || uniqueMaterials.has(name)) {



            return;



          }



          const fallback = target === "material" ? layer.materialPropsFallback : layer.fillMaterialPropsFallback;



          const materialNode = doc.createElement("Material");



          materialNode.setAttribute("Name", name);



          appendDoubleElement(doc, materialNode, "Permittivity", fallback.permittivity);



          appendDoubleElement(doc, materialNode, "DielectricLossTangent", fallback.lossTangent);



          appendDoubleElement(doc, materialNode, "Conductivity", fallback.conductivity);



          materials.appendChild(materialNode);



          uniqueMaterials.set(name, true);



        });



      });







      const layers = doc.createElement("Layers");



      if (state.lengthUnit.trim()) {



        layers.setAttribute("LengthUnit", state.lengthUnit.trim());



      }



      stackup.appendChild(layers);







      state.layers.forEach((layer) => {



        const layerNode = doc.createElement("Layer");



        setAttributeIfValue(layerNode, "Name", layer.name);



        setAttributeIfValue(layerNode, "Type", layer.type);



        setAttributeIfValue(layerNode, "Material", layer.material);



        setAttributeIfValue(layerNode, "FillMaterial", layer.fillMaterial);



        setAttributeIfValue(layerNode, "Thickness", layer.thickness);



        setAttributeIfValue(layerNode, "EtchFactor", layer.etchFactor);



        setAttributeIfValue(layerNode, "Color", layer.color);







        appendHuray(doc, layerNode, "HuraySurfaceRoughness", layer.hurayTopRatio, layer.hurayTopRadius);



        appendHuray(doc, layerNode, "HurayBottomSurfaceRoughness", layer.hurayBottomRatio, layer.hurayBottomRadius);



        appendHuray(doc, layerNode, "HuraySideSurfaceRoughness", layer.huraySideRatio, layer.huraySideRadius);







        layers.appendChild(layerNode);



      });







      const serializer = new XMLSerializer();



      const raw = serializer.serializeToString(doc);



      const formatted = formatXml(raw.replace(/^<\?xml[^>]*>/, ""));



      return '<?xml version="1.0" encoding="UTF-8" standalone="no"?>\n' + formatted;



    }







    function appendDoubleElement(doc, parent, tagName, value) {



      const trimmed = value && value.toString ? value.toString().trim() : "";



      if (!trimmed) {



        return;



      }



      const container = doc.createElement(tagName);



      const doubleNode = doc.createElement("Double");



      doubleNode.textContent = trimmed;



      container.appendChild(doubleNode);



      parent.appendChild(container);



    }







    function setAttributeIfValue(node, attribute, value) {



      const trimmed = value && value.toString ? value.toString().trim() : "";



      if (!trimmed) {



        return;



      }



      node.setAttribute(attribute, trimmed);



    }







    function appendHuray(doc, parent, tagName, ratio, radius) {



      const ratioValue = ratio && ratio.toString ? ratio.toString().trim() : "";



      const radiusValue = radius && radius.toString ? radius.toString().trim() : "";



      if (!ratioValue && !radiusValue) {



        return;



      }



      const node = doc.createElement(tagName);



      if (ratioValue) {



        node.setAttribute("HallHuraySurfaceRatio", ratioValue);



      }



      if (radiusValue) {



        node.setAttribute("NoduleRadius", radiusValue);



      }



      parent.appendChild(node);



    }







    function formatXml(xml) {



      const padding = "  ";



      let formatted = "";



      let pad = 0;



      xml



        .replace(/>\s+</g, "><")



        .split(/(?=<)/g)



        .map((node) => node.trim())



        .filter((node) => node.length > 0)



        .forEach((node) => {



          if (node.indexOf("</") === 0) {



            pad = Math.max(pad - 1, 0);



          }



          formatted += padding.repeat(pad) + node + "\n";



          if (node.indexOf("</") !== 0 && node.indexOf("/>") !== node.length - 2 && node.indexOf("</") === -1) {



            pad += 1;



          }



        });



      return formatted.trim();



    }







    function findChildElement(parent, tagName) {



      if (!parent) {



        return null;



      }



      const nodes = parent.children || [];



      for (let i = 0; i < nodes.length; i += 1) {



        const child = nodes[i];



        if (!tagName || child.localName === tagName) {



          return child;



        }



      }



      return null;



    }







    function findChildElements(parent, tagName) {



      const result = [];



      if (!parent) {



        return result;



      }



      const nodes = parent.children || [];



      for (let i = 0; i < nodes.length; i += 1) {



        const child = nodes[i];



        if (!tagName || child.localName === tagName) {



          result.push(child);



        }



      }



      return result;



    }



  