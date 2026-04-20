/**
 * Returns system cost indices.
 *
 * @param {number} systemId (optional) Filter the list to a specific system.
 * @customfunction
 */
function costIndicesBySystem(systemId=31000376) {
  const data = GESI.invokeRaw('industry_systems');

  const indiciesBySystem = {};

  data.forEach((obj) => {
    const activityCostMap = {};

    obj.cost_indices.forEach((costIndex) => activityCostMap[costIndex.activity] = costIndex.cost_index);

    indiciesBySystem[obj.solar_system_id] = activityCostMap;
  });
console.log("Indicies by system: "+indiciesBySystem);
  const result = [
    ['solar_system_id', 'manufacturing', 'researching_time_efficiency', 'researching_material_efficiency', 'copying', 'invention', 'reaction'], 
  ];

  if (systemId) {
    const systemInfo = indiciesBySystem[systemId];

    result.push([
      systemId, systemInfo.manufacturing, systemInfo.researching_time_efficiency, systemInfo.researching_material_efficiency, systemInfo.copying, systemInfo.invention, systemInfo.reaction
    ]);

    return result;
  }

  for (const [systemId, systemInfo] of Object.entries(indiciesBySystem)) {
    result.push([
      systemId, systemInfo.manufacturing, systemInfo.researching_time_efficiency, systemInfo.researching_material_efficiency, systemInfo.copying, systemInfo.invention, systemInfo.reaction
    ]);
  }
  console.log("result: "+result)
  return result;
}