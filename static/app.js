function pad(n) {
  return String(n).padStart(2, "0");
}

function nowTime() {
  const d = new Date();
  const minutes = Math.round(d.getMinutes() / 5) * 5;
  d.setMinutes(minutes);
  return `${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

function durationHours(start, end) {
  if (!start || !end) return 0;
  const [sh, sm] = start.split(":").map(Number);
  const [eh, em] = end.split(":").map(Number);
  let ini = sh * 60 + sm;
  let fin = eh * 60 + em;
  if (fin < ini) fin += 24 * 60;
  return Math.max(0, Math.round(((fin - ini) / 60) * 100) / 100);
}

function lunchDiscountHours(start, end) {
  const raw = durationHours(start, end);
  if (raw <= 1 || !start || !end) return 0;
  const [sh, sm] = start.split(":").map(Number);
  const [eh, em] = end.split(":").map(Number);
  let ini = sh * 60 + sm;
  let fin = eh * 60 + em;
  if (fin < ini) fin += 24 * 60;
  const lunchStart = 12 * 60;
  const lunchEnd = 16 * 60;
  return ini < lunchEnd && fin > lunchStart ? 1 : 0;
}

function effectiveDurationHours(start, end) {
  return Math.max(0, Math.round((durationHours(start, end) - lunchDiscountHours(start, end)) * 100) / 100);
}

function selectedWorkerInputs() {
  return Array.from(document.querySelectorAll("input[name='trabajadores_asignados']:checked"));
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function updateWorkerDropdown() {
  const selected = selectedWorkerInputs();
  const triggerText = document.getElementById("workerTriggerText");
  const badge = document.getElementById("workerBadge");
  const note = document.getElementById("workerLimitNote");
  if (triggerText) {
    triggerText.textContent = selected.length
      ? selected.map((input) => input.value.split(" ").slice(0, 2).join(" ")).join(", ")
      : "Seleccionar personas";
  }
  if (badge) badge.textContent = `${selected.length}/8`;
  if (note && selected.length <= 8) note.textContent = "Selecciona hasta 8 personas.";
}

function updateHHLimitAlert(extraHours) {
  const alertBox = document.getElementById("hhLimitAlert");
  if (!alertBox) return;
  const data = document.getElementById("hhPersonasData");
  let current = {};
  if (data) {
    try {
      current = JSON.parse(data.textContent || "{}");
    } catch (error) {
      current = {};
    }
  }
  const excedidos = selectedWorkerInputs()
    .map((input) => {
      const acumuladas = Number(current[input.value] || 0);
      const total = Math.round((acumuladas + extraHours) * 100) / 100;
      return { nombre: input.value, acumuladas, total };
    })
    .filter((item) => item.total > 11);

  if (!excedidos.length) {
    alertBox.hidden = true;
    alertBox.innerHTML = "";
    return;
  }

  alertBox.hidden = false;
  const detalle = excedidos
    .map((item) => `${escapeHtml(item.nombre)}: quedaria con ${item.total} HH`)
    .join("<br>");
  alertBox.innerHTML = `<strong>Control interno:</strong> revisa la asignacion antes de guardar.<br>${detalle}`;
}

function classifyText(estado, causal, hh) {
  const code = (causal || "").slice(0, 3);
  if (estado === "Ejecutado" || estado === "En proceso") return `${hh} HH directas`;
  if (estado === "Parcial") return "Parcial: separar HH directas/indirectas si corresponde";
  if (estado === "No ejecutado") return `${hh} HH no utilizadas`;
  if (estado === "Stand-by") {
    if (["C02", "C03", "C04", "C05", "C06", "C08"].includes(code)) return `${hh} HH indirectas por condicion operacional`;
    return `${hh} HH no utilizadas`;
  }
  return "Clasificacion segun estado y causal.";
}

function updateSummary() {
  const form = document.getElementById("frenteForm");
  if (!form) return;
  const start = form.hora_inicio.value;
  const end = form.hora_termino.value;
  const dur = durationHours(start, end);
  const colacion = lunchDiscountHours(start, end);
  const efectiva = effectiveDurationHours(start, end);
  const estado = form.estado.value;
  const causal = form.causal.value;
  const selectedWorkers = selectedWorkerInputs().length;
  const hh = Math.round(efectiva * selectedWorkers * 100) / 100;
  document.getElementById("sumDuracion").textContent = `${dur} h`;
  document.getElementById("sumColacion").textContent = `${colacion} h`;
  document.getElementById("sumEfectiva").textContent = `${efectiva} h`;
  document.getElementById("sumPersonas").textContent = String(selectedWorkers);
  document.getElementById("sumHH").textContent = selectedWorkers ? `${hh} HH` : "0 HH";
  document.getElementById("sumClasificacion").textContent = classifyText(estado, causal, "Total");
  document.getElementById("partialBox").classList.toggle("show", estado === "Parcial");
  updateWorkerDropdown();
  updateHHLimitAlert(efectiva);
}

function fillSelect(id, value) {
  const el = document.getElementById(id);
  if (!el || value == null) return;
  const exists = Array.from(el.options || []).some((o) => o.value === value);
  if (exists || el.tagName !== "SELECT") el.value = value;
}

function shiftHours(turno) {
  return turno === "Noche"
    ? [19, 20, 21, 22, 23, 0, 1, 2, 3, 4, 5, 6, 7]
    : [7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19];
}

function setQuickTime(form, value) {
  const target = form.hora_inicio.value ? form.hora_termino : form.hora_inicio;
  target.value = value;
  updateSummary();
}

function renderTimeChips(form) {
  const container = document.getElementById("timechips");
  if (!container) return;
  container.innerHTML = "";
  shiftHours(form.turno.value).forEach((hour) => {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.dataset.time = `${pad(hour)}:00`;
    btn.textContent = pad(hour);
    btn.addEventListener("click", () => setQuickTime(form, btn.dataset.time));
    container.appendChild(btn);
  });
  const nowBtn = document.createElement("button");
  nowBtn.type = "button";
  nowBtn.dataset.now = "true";
  nowBtn.textContent = "Ahora";
  nowBtn.addEventListener("click", () => setQuickTime(form, nowTime()));
  container.appendChild(nowBtn);
}

function readJsonData(id, fallback) {
  const data = document.getElementById(id);
  if (!data) return fallback;
  try {
    return JSON.parse(data.textContent || "null") ?? fallback;
  } catch (error) {
    return fallback;
  }
}

function applyFrenteShiftHours(form) {
  if (!form.hora_inicio || !form.hora_termino || !form.turno) return;
  if (form.turno.value === "Noche") {
    form.hora_inicio.value = "19:15";
    form.hora_termino.value = "07:15";
  } else {
    form.hora_inicio.value = "07:15";
    form.hora_termino.value = "19:15";
  }
  updateSummary();
}

function initFrentes() {
  const form = document.getElementById("frenteForm");
  if (!form) return;
  const fecha = document.getElementById("fecha");
  if (fecha && !fecha.value) fecha.value = new Date().toISOString().slice(0, 10);

  form.addEventListener("input", updateSummary);
  form.addEventListener("change", updateSummary);
  form.addEventListener("submit", (event) => {
    const alertBox = document.getElementById("hhLimitAlert");
    if (alertBox && !alertBox.hidden) {
      event.preventDefault();
      alertBox.scrollIntoView({ behavior: "smooth", block: "center" });
    }
  });

  renderTimeChips(form);
  if (!form.hora_inicio.value && !form.hora_termino.value) applyFrenteShiftHours(form);
  form.turno.addEventListener("change", () => {
    renderTimeChips(form);
    applyFrenteShiftHours(form);
  });

  document.querySelectorAll("[data-action]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const action = btn.dataset.action;
      if (action === "start") {
        form.estado.value = "En proceso";
        form.causal.value = "C01 Sin desviacion";
        form.hora_inicio.value = nowTime();
        form.hora_termino.value = "";
      }
      if (action === "close") {
        form.estado.value = "Ejecutado";
        form.causal.value = "C01 Sin desviacion";
        if (!form.hora_inicio.value) form.hora_inicio.value = nowTime();
        form.hora_termino.value = nowTime();
      }
      if (action === "noexec") {
        form.estado.value = "No ejecutado";
        if (!form.hora_inicio.value) form.hora_inicio.value = nowTime();
        form.hora_termino.value = form.hora_inicio.value;
        form.causal.value = "C06 Falta de frente disponible";
      }
      const lastFrente = readJsonData("lastFrenteData", null);
      if (action === "duplicate" && lastFrente) {
        ["supervisor", "nombre_tarea", "equipo", "actividad", "estado", "causal", "hora_inicio", "hora_termino"].forEach((key) => {
          fillSelect(key, lastFrente[key]);
        });
        form.observacion.value = "";
        const selected = lastFrente.trabajadores_asignados || "";
        document.querySelectorAll("input[name='trabajadores_asignados']").forEach((input) => {
          input.checked = selected.includes(input.value);
        });
      }
      updateSummary();
    });
  });
  const workerDropdown = document.getElementById("workerDropdown");
  const workerTrigger = document.getElementById("workerTrigger");
  const workerMenu = document.getElementById("workerMenu");
  const clearWorkers = document.getElementById("clearWorkers");
  const selectAllWorkers = document.getElementById("selectAllWorkers");
  if (workerTrigger && workerMenu) {
    workerTrigger.addEventListener("click", () => {
      workerDropdown.classList.toggle("open");
    });
    document.addEventListener("click", (event) => {
      if (!workerDropdown.contains(event.target)) workerDropdown.classList.remove("open");
    });
  }
  document.querySelectorAll("input[name='trabajadores_asignados']").forEach((input) => {
    input.addEventListener("change", () => {
      const selected = selectedWorkerInputs();
      if (selected.length > 8) {
        input.checked = false;
        const note = document.getElementById("workerLimitNote");
        if (note) note.textContent = "Maximo 8 personas por turno.";
      }
      updateSummary();
    });
  });
  if (clearWorkers) {
    clearWorkers.addEventListener("click", () => {
      document.querySelectorAll("input[name='trabajadores_asignados']").forEach((input) => {
        input.checked = false;
      });
      updateSummary();
    });
  }
  if (selectAllWorkers) {
    selectAllWorkers.addEventListener("click", () => {
      document.querySelectorAll("input[name='trabajadores_asignados']").forEach((input, index) => {
        input.checked = index < 8;
      });
      const note = document.getElementById("workerLimitNote");
      if (note) note.textContent = "Todo el personal disponible seleccionado.";
      updateSummary();
    });
  }
  document.querySelectorAll("[data-workers]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const checked = btn.dataset.workers === "all";
      document.querySelectorAll("input[name='trabajadores_asignados']").forEach((input) => {
        input.checked = checked;
      });
      updateSummary();
    });
  });
  updateSummary();
}

function initAttendance() {
  const form = document.getElementById("attendanceForm");
  if (!form) return;
  const turnoSelect = document.getElementById("turnoSelect");
  const horaIngreso = document.getElementById("hora_ingreso");
  const horaSalida = document.getElementById("hora_salida");
  const count = document.getElementById("attendanceCount");
  const badge = document.getElementById("attendanceBadge");
  const triggerText = document.getElementById("attendanceTriggerText");
  const dropdown = document.getElementById("attendanceDropdown");
  const trigger = document.getElementById("attendanceTrigger");
  const clear = document.getElementById("clearAttendance");
  const selectAll = document.getElementById("selectAllAttendance");
  const inputs = Array.from(document.querySelectorAll("input[name='trabajadores_presentes']"));
  const absentDropdown = document.getElementById("absentDropdown");
  const absentTrigger = document.getElementById("absentTrigger");
  const absentCount = document.getElementById("absentCount");
  const absentBadge = document.getElementById("absentBadge");
  const absentTriggerText = document.getElementById("absentTriggerText");
  const clearAbsent = document.getElementById("clearAbsent");
  const absentInputs = Array.from(document.querySelectorAll("input[name='trabajadores_ausentes']"));
  const presentByName = new Map(inputs.map((input) => [input.value, input]));
  const absentByName = new Map(absentInputs.map((input) => [input.dataset.workerName, input]));
  const applyShiftHours = () => {
    if (!turnoSelect || !horaIngreso || !horaSalida) return;
    if (turnoSelect.value === "Noche") {
      horaIngreso.value = "19:15";
      horaSalida.value = "07:15";
    } else {
      horaIngreso.value = "07:15";
      horaSalida.value = "19:15";
    }
  };
  if (turnoSelect) {
    turnoSelect.addEventListener("change", applyShiftHours);
    applyShiftHours();
  }
  const update = () => {
    const selectedInputs = inputs.filter((input) => input.checked);
    const selected = selectedInputs.length;
    const absentSelected = absentInputs.filter((input) => input.checked);
    if (count) count.textContent = selected === 1 ? "1 seleccionado" : `${selected} seleccionados`;
    if (badge) badge.textContent = String(selected);
    if (triggerText) {
      triggerText.textContent = selected
        ? selectedInputs.map((input) => input.value.split(" ").slice(0, 2).join(" ")).join(", ")
        : "Seleccionar soldadores presentes";
    }
    if (absentCount) absentCount.textContent = absentSelected.length === 1 ? "1 ausente" : `${absentSelected.length} ausentes`;
    if (absentBadge) absentBadge.textContent = String(absentSelected.length);
    if (absentTriggerText) {
      absentTriggerText.textContent = absentSelected.length
        ? absentSelected.map((input) => input.dataset.workerName.split(" ").slice(0, 2).join(" ")).join(", ")
        : "Seleccionar ausentes";
    }
    absentInputs.forEach((input) => {
      const select = input.closest(".absent-check")?.querySelector("select");
      if (select) select.disabled = !input.checked;
    });
  };
  if (trigger && dropdown) {
    trigger.addEventListener("click", () => {
      dropdown.classList.toggle("open");
    });
    document.addEventListener("click", (event) => {
      if (!dropdown.contains(event.target)) dropdown.classList.remove("open");
    });
  }
  if (absentTrigger && absentDropdown) {
    absentTrigger.addEventListener("click", () => {
      absentDropdown.classList.toggle("open");
    });
    document.addEventListener("click", (event) => {
      if (!absentDropdown.contains(event.target)) absentDropdown.classList.remove("open");
    });
  }
  inputs.forEach((input) => {
    input.addEventListener("change", () => {
      if (input.checked) {
        const absent = absentByName.get(input.value);
        if (absent) absent.checked = false;
      }
      update();
    });
  });
  absentInputs.forEach((input) => {
    input.addEventListener("change", () => {
      if (input.checked) {
        const present = presentByName.get(input.dataset.workerName);
        if (present) present.checked = false;
      }
      update();
    });
  });
  if (clear) {
    clear.addEventListener("click", () => {
      inputs.forEach((input) => {
        input.checked = false;
      });
      update();
    });
  }
  if (selectAll) {
    selectAll.addEventListener("click", () => {
      inputs.forEach((input) => {
        input.checked = true;
      });
      absentInputs.forEach((input) => {
        input.checked = false;
      });
      update();
    });
  }
  if (clearAbsent) {
    clearAbsent.addEventListener("click", () => {
      absentInputs.forEach((input) => {
        input.checked = false;
      });
      update();
    });
  }
  update();
}

document.addEventListener("DOMContentLoaded", initFrentes);
document.addEventListener("DOMContentLoaded", initAttendance);
