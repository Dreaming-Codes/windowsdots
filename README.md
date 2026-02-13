```batch
winget install --id Git.Git --exact --silent --disable-interactivity --accept-package-agreements --accept-source-agreements --scope user && winget install --id twpayne.chezmoi --exact --silent --disable-interactivity --accept-package-agreements --accept-source-agreements --scope user && chezmoi init --apply git@github.com:Dreaming-Codes/windowsdots
```
