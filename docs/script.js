(function () {
  "use strict";

  var translations = {
    en: {
      "hero.tagline":
        "Render Zsigmondy-Palmer dental notation as images.",
      "hero.download": "Download",
      "hero.github": "View on GitHub",
      "about.title": "About",
      "about.text":
        "palmer-type compiles Zsigmondy-Palmer dental notation into publication-ready PNG, JPEG, or PDF images. Enter tooth numbers, click Render, and copy or save the result. No TeX expertise required.",
      "about.text2": "",
      "features.title": "Features",
      "features.gui.title": "GUI Rendering",
      "features.gui.desc":
        "Enter notation in a visual cross layout and render crisp images in seconds.",
      "features.cli.title": "CLI & Batch",
      "features.cli.desc":
        "Automate rendering from the command line with single or batch JSON input.",
      "features.converter.title": "Word Converter",
      "features.converter.desc":
        "Replace \\Palmer commands in .docx files with rendered images automatically.",
      "screenshot.title": "Screenshot",
      "screenshot.alt":
        "palmer-type GUI window with rendered dental notation preview",
      "download.title": "Download",
      "download.text":
        "Several executables are available on the GitHub Releases page. The installer (palmer-type-\u2026-win-x64-setup.exe) is recommended for most users.",
      "download.button": "Download Latest Release",
      "download.note_summary": "Note on first run",
      "download.note_text":
        "The initial launch requires an internet connection. Tectonic automatically downloads the TeX support files it needs (~100\u00a0MB) and caches them locally. Subsequent launches work offline.",
      "related.title": "Related Projects",
      "related.sty_desc":
        "The underlying LaTeX package for Zsigmondy-Palmer dental notation. Licensed under LPPL 1.3+.",
      "related.repo_desc":
        "Source code and issue tracker for this tool.",
      "footer.grant":
        "Supported by JSPS KAKENHI Grant Number JP25K15395.",
      "donate.title": "Support This Project",
      "donate.text":
        "palmer-type is free and open-source. Donations help fund ongoing research and development of this tool and the underlying palmer.sty package. If you find this project useful in your research, clinical work, or publishing, your support is greatly appreciated.",
    },
    ja: {
      "hero.tagline":
        "Zsigmondy-Palmer式歯式記号を画像保存できるアプリです。",
      "hero.download": "ダウンロード",
      "hero.github": "GitHubで見る",
      "about.title": "概要",
      "about.text":
        "palmer-typeは、Zsigmondy-Palmer式歯式記号をPNG・JPEG・PDF画像にコンパイルします。歯番号を入力し、Renderをクリックするだけで、出版品質の画像をコピー・保存できます。",
      "about.text2": "",
      "features.title": "機能",
      "features.gui.title": "GUIレンダリング",
      "features.gui.desc":
        "十字型レイアウトに歯式を入力し、数秒で鮮明な画像を生成します。",
      "features.cli.title": "CLI・バッチ処理",
      "features.cli.desc":
        "コマンドラインから単体またはJSON一括入力で自動レンダリング。",
      "features.converter.title": "Wordコンバーター",
      "features.converter.desc":
        ".docxファイル内の\\Palmerコマンドを自動的にレンダリング画像に置換します。",
      "screenshot.title": "スクリーンショット",
      "screenshot.alt":
        "palmer-type GUIウィンドウ — 歯式記号レンダリングプレビュー",
      "download.title": "ダウンロード",
      "download.text":
        "GitHub Releasesページに複数の実行ファイルがあります。インストーラー版（palmer-type-\u2026-win-x64-setup.exe）の利用を推奨します。",
      "download.button": "最新版をダウンロード",
      "download.note_summary": "初回起動時の注意",
      "download.note_text":
        "初回起動時にはインターネット接続が必要です。Tectonicが必要なTeXサポートファイル（約100 MB）を自動ダウンロードしてローカルにキャッシュします。2回目以降はオフラインで動作します。",
      "related.title": "関連プロジェクト",
      "related.sty_desc":
        "Zsigmondy-Palmer式歯式記号のためのLaTeXパッケージ。LPPL 1.3+ライセンス。",
      "related.repo_desc":
        "本ツールのソースコードとイシュートラッカー。",
      "footer.grant":
        "本研究はJSPS科研費 JP25K15395 の助成を受けたものです。",
      "donate.title": "寄付",
      "donate.text":
        "palmer-typeはフリーかつオープンソースです。研究活動を継続するためのご支援として、研究助成金のご寄付を随時受け付けております。寄付金控除の対象となりますので、領収書等が必要な方は事前にご連絡ください。",
    },
  };

  // about.text2 contains HTML with a link — handle separately
  var htmlTranslations = {
    en: {
      "about.text2":
        'Built on <a href="https://github.com/yosukey/palmer-latex">palmer.sty</a>, a LaTeX package for Palmer notation, and bundled with Tectonic as its TeX engine \u2014 no TeX distribution needed.',
    },
    ja: {
      "about.text2":
        'Palmer記法のためのLaTeXパッケージ <a href="https://github.com/yosukey/palmer-latex">palmer.sty</a> をベースに、TeXエンジンTectonicを同梱。TeX環境のインストールは不要です。',
    },
  };

  function setLanguage(lang) {
    if (!translations[lang]) return;

    // Update text content
    document.querySelectorAll("[data-i18n]").forEach(function (el) {
      var key = el.getAttribute("data-i18n");

      // Check if this key has HTML translation
      if (htmlTranslations[lang] && htmlTranslations[lang][key] !== undefined) {
        el.innerHTML = htmlTranslations[lang][key];
      } else if (translations[lang][key] !== undefined) {
        el.textContent = translations[lang][key];
      }
    });

    // Update alt attributes
    document.querySelectorAll("[data-i18n-alt]").forEach(function (el) {
      var key = el.getAttribute("data-i18n-alt");
      if (translations[lang][key] !== undefined) {
        el.alt = translations[lang][key];
      }
    });

    // Update <html lang>
    document.documentElement.lang = lang;

    // Update toggle buttons
    document.querySelectorAll(".lang-btn").forEach(function (btn) {
      btn.classList.toggle("lang-active", btn.getAttribute("data-lang") === lang);
    });

    // Persist
    try {
      localStorage.setItem("palmer-type-lang", lang);
    } catch (e) {
      // localStorage unavailable
    }
  }

  // Init
  document.addEventListener("DOMContentLoaded", function () {
    // Bind language toggle
    document.querySelectorAll(".lang-btn").forEach(function (btn) {
      btn.addEventListener("click", function () {
        setLanguage(btn.getAttribute("data-lang"));
      });
    });

    // Restore saved language
    try {
      var saved = localStorage.getItem("palmer-type-lang");
      if (saved && translations[saved]) {
        setLanguage(saved);
      }
    } catch (e) {
      // localStorage unavailable
    }

    // Fetch latest release version from GitHub API
    fetch("https://api.github.com/repos/yosukey/palmer-type/releases/latest")
      .then(function (r) { return r.ok ? r.json() : null; })
      .then(function (data) {
        if (data && data.tag_name) {
          var badge = document.getElementById("version-badge");
          if (badge) {
            badge.textContent = data.tag_name;
            badge.style.display = "";
          }
        }
      })
      .catch(function () { /* silently hide badge if offline or no releases */ });
  });
})();
