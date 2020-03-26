//prefers-color-scheme : light or dark
//themes: generic.light or generic.dark
if (window.matchMedia && window.matchMedia('(prefers-color-scheme: light)').matches) {
    DevExpress.ui.themes.current("generic.light");
}