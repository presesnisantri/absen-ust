export function parseClass(classString) {
  if (!classString) return { grade: 999, major: "", variant: "" };
  
  const match = classString.trim().match(/^(\d+)\s*([A-Za-z]+)?(.*)?$/);
  if (match) {
    const grade = parseInt(match[1], 10) || 999;
    const major = (match[2] || "").trim().toUpperCase();
    const variant = (match[3] || "").trim();
    return { grade, major, variant };
  }
  return { grade: 999, major: classString.trim(), variant: "" };
}

export function sortClasses(a, b) {
  const classA = a.kelas || "";
  const classB = b.kelas || "";
  
  const parsedA = parseClass(classA);
  const parsedB = parseClass(classB);

  if (parsedA.grade !== parsedB.grade) {
    return parsedA.grade - parsedB.grade;
  }
  
  if (parsedA.major !== parsedB.major) {
    return parsedA.major.localeCompare(parsedB.major);
  }
  
  if (parsedA.variant !== parsedB.variant) {
    return parsedA.variant.localeCompare(parsedB.variant);
  }
  
  const nameA = a.nama_guru || "";
  const nameB = b.nama_guru || "";
  return nameA.localeCompare(nameB);
}
