# figtree
Growing figs on the tree of knowledge

```
python3 download_from_gdrive.py
python3 build_weekly_ppt.py >> temp_output.txt
python3 upload_to_gdrive.py
```

put this one in travis-ci

```
CREDENTIALS='{"access_token": ...'
```

run this before you run the download cmd
```
echo $CREDENTIALS >> mycreds1.txt 
```