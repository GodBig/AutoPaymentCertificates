$(function () {
    function getCookie(name) {
        let r = document.cookie.match("\\b" + name + "=([^;]*)\\b");
        return r ? r[1] : undefined;
    }

    let csrf = getCookie("_xsrf");
    let flow = new Flow({
        target: '/input',
        query: {'_xsrf': csrf}
    });
    // Flow.js isn't supported, fall back on a different method
    if (!flow.support) location.href = '/some-old-crappy-uploader';
    flow.assignBrowse(document.getElementById('lpo_xlsx'));
    flow.assignBrowse(document.getElementById('lpo_pdf'));
    flow.assignBrowse(document.getElementById('tax_invoice'));
    flow.assignDrop(document.getElementById('reset'));
})
