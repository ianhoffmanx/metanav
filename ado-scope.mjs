// WIQL literals and team sprint scope are kept separate for regression checks.
export const wiqlLiteral = value => String(value).replace(/'/g, "''");
export function projectScope(project) { return project ? ` AND [System.TeamProject] = '${wiqlLiteral(decodeURIComponent(project))}'` : ''; }
export function iterationScope(iterations) { return iterations.length ? ` AND [System.IterationPath] IN (${iterations.map(i=>`'${wiqlLiteral(i.path)}'`).join(',')})` : ''; }
