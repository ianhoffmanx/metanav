// Opt-in marketing scope; the original collector shapes remain available to other users.
const contentFile = /\.(?:md|mdx|markdown|rst|txt)$/i;
const contentAsset = /(?:^|\/)(?:content|posts|blog|blogs|news|public\/uploads)\/.*\.(?:json|ya?ml|html|png|jpe?g|webp|svg|gif)$/i;
const technicalTitle = /\b(?:refactor|tailwind|sitemap|static rendering|content-neutral)\b/i;
const editorialTitle = /\b(?:blog|marketing|social|campaign|newsletter|copy)\b/i;
const contentWords = /\b(?:blog|content|copy|marketing|social|campaign|newsletter|email assets|release post|landing page)\b/i;
export function inGithubScope(url, options) {
  if (!options || !url) return true;
  let u; try { u = new URL(url); } catch { return false; }
  if (u.hostname !== 'github.com') return false;
  const repo = u.pathname.split('/').slice(1,3).join('/').toLowerCase();
  return options.orgs.some(org=>repo.startsWith(org.toLowerCase()+'/')) && !(options.exclude_repos || []).some(x=>repo===x.toLowerCase() || repo.endsWith('/'+x.toLowerCase()));
}
export function contentRelevant(item, options) {
  if (!options) return true;
  if (!inGithubScope(item.url, options)) return false;
  const repo = item.url.split('/').slice(3,5).join('/').toLowerCase();
  const configured = options.content_repos.some(x=>x.toLowerCase()===repo);
  const paths = (item.files?.nodes || item.files || []).map(f=>typeof f==='string'?f:f.path);
  // A content repo can also contain application code. Require content evidence.
  if (technicalTitle.test(item.title || '') && !editorialTitle.test(item.title || '')) return false;
  return (configured && paths.some(p=>contentFile.test(p) || contentAsset.test(p))) || contentWords.test(item.title || '');
}
export function marketingGithub(data, options) {
  if (!options || !data?.ok) return data;
  const presentation = item => {
    const { size, ci, unresolved_threads, files, ...rest } = item;
    return { ...rest, content_review:true, ...(item.kind === 'pbi' ? {} : ci==='FAILURE'||ci==='ERROR' ? {publication_blocked:true,blocker:'Publication check failed; ask the author to fix it.'}:{}),
      ...(item.prs && {prs:item.prs.filter(p=>contentRelevant(p,options)).map(presentation)}) };
  };
  return {...data, excluded_prs:[...(data.work||[]),...(data.review_requests||[])].filter(x=>!contentRelevant(x,options)).map(x=>x.url),
    work:(data.work||[]).filter(w=>contentRelevant(w,options)).map(presentation),
    review_requests:(data.review_requests||[]).filter(w=>contentRelevant(w,options)).map(presentation),
    mentions:(data.mentions||[]).filter(w=>inGithubScope(w.url,options)),
  };
}
// Never carry excluded side projects into another judge prompt, including old panel data.
export function marketingPrevious(previous, options, currentGithub) {
  if (!options || !previous) return previous;
  return {...previous,...Object.fromEntries(['queue','waiting','highlights','done','work','review_requests'].map(key=>[key,(previous[key]||[]).filter(x=>{
    if ((x.source==='github'||key==='work'||key==='review_requests') && currentGithub?.ok && currentGithub.excluded_prs?.includes(x.url)) return false;
    const urls=JSON.stringify(x).match(/https:\/\/github\.com\/[^\s"?#]+/g)||[];
    return urls.every(u=>inGithubScope(u,options));
  })]))};
}
