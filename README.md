# iitmigration
A series of scripts written in Python for data manipulation in Excel using openpyxl

Images in Time is a legacy image database created by the Toledo Lucas County Public Library
http://images2.toledolibrary.org/image_dc_new.asp

These images are being uploaded to CONTENTdm, through Ohio Memory
https://www.ohiomemory.org/digital/collection/p16007coll33/

Metadata for the records stored in Images in Time was exported by the excellent Roxanna Foster

This metadata still required a great deal of clean-up so this repo includes scripts to automate as much of the process as possible

The scripts are written in Python making extensive use of the openpyxl library as the metadata is stored in xlsx spreadsheets

Some scipts also utilize regex

Each folder in the repo contains the initital data copied over from initial export as well as the final metadata sheet that was used to add images and metadata to TLCPL CONTENTdm collections and the corresponding Python scripts used to manipulate that metadata

Progress

2020-12-09 | 892 images | John Vanderlip Photograph Collection

2020-12-11 | 830 images | Herral Long Photograph Collection

2020-12-16 | 617 images | Charles R. Mensing Photograph Collection

2020-12-29 | 242 images | Korb Photographic Company Collection

2020-12-30 | 764 images | Milton Zink Photograph Collection

2020-12-31 | 218 images | Rudolph Gartner Photograph Collection
