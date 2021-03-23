# iitmigration
A series of scripts written in Python for data manipulation in Excel using openpyxl

Images in Time is a legacy image database created by the Toledo Lucas County Public Library
http://images2.toledolibrary.org/image_dc_new.asp

These images are being uploaded to CONTENTdm, through Ohio Memory
https://www.ohiomemory.org/digital/collection/p16007coll33/

Metadata for the records stored in Images in Time was exported by the excellent Roxanna Foster

This metadata still required a great deal of clean-up so this repo includes scripts to automate as much of the process as possible

The scripts are written in Python making extensive use of the openpyxl library as the metadata is stored in xlsx spreadsheets

Some scipts also utilize regex. While I always attempted to provide before and after views of the metadata, this was occasionally forgotten and only the final product is available

Each folder in the repo contains the initital data copied over from initial export as well as the final metadata sheet that was used to add images and metadata to TLCPL CONTENTdm collections and the corresponding Python scripts used to manipulate that metadata

Progress

2020-12-09 | 892 images | John Vanderlip Photograph Collection | aalh_iit_vanderlipcollection

2020-12-11 | 830 images | Herral Long Photograph Collection | aalh_iit_herrallongcollection

2020-12-16 | 617 images | Charles R. Mensing Photograph Collection | aalh_iit_charlesmensingcollection

2020-12-29 | 242 images | Korb Photographic Company Collection | aalh_iit_korbphotographiccompany

2020-12-30 | 764 images | Milton Zink Photograph Collection | aalh_iit_miltonzinkcollection

2020-12-31 | 218 images | Rudolph Gartner Photograph Collection | aalh_iit_rudolphgartnercollection

2020-12-31 | 136 images | Wilbur Hague Photograph Collection | aalh_iit_wilburhaguecollection

2021-01-06 | 300 images | Hauger Photographic Corporation Collection | aalh_iit_haugerphotocorp

2021-01-07 | 215 images | Howard MacKenzie Photograph Collection | aalh_iit_howardmackenziecollection

2021-01-07 | 194 images | J. Doyle Witgen Photograph Collection | aalh_iit_jdoylewitgencollection

2021-01-21 | 2636 images | Ted Ligibel Photograph Collection | aalh_iit_tedligibelcollection

2021-01-26 | 498 images | aalh_iit_buildings_01

2021-01-27 | 495 images | aalh_iit_peopleportraits_001

2021-01-29 | 507 images | aalh_iit_transportation_001

2021-01-30 | 508 images | aalh_iit_peopleportraits_002

2021-02-02 | 499 images | aalh_iit_buildings_02

2021-02-04 | 498 images | aalh_iit_peopleportraits_003

2021-02-04 | 493 images | aalh_iit_buildings_03

2021-02-05 | 493 images | aalh_iit_peopleportraits_004

2021-02-09 | 631 images | aalh_iit_transportation_002

2021-02-11 | 523 images | aalh_iit_buildings_04

2021-02-17 | 532 images | aalh_iit_peopleportraits_005

2021-02-18 | 581 images | aalh_iit_peopleportraits_006

2021-02-23 | 491 images | aalh_iit_buildings_005

2021-02-23 | 517 images | aalh_iit_peopleportraits_007

2021-03-04 | 502 images | aalh_iit_transportation_003

2021-03-05 | 505 images | aalh_iit_buildings_006

2021-03-08 | 539 images | aalh_iit_peopleportraits_008

2021-03-09 | 499 images | aalh_iit_buildings_007

2021-03-11 | 259 images | aalh_iit_tlcpl_001

2021-03-11 | 595 images | aalh_iit_churches_001

2021-03-13 | 493 images | aalh_iit_peopleportraits_009

2021-03-13 | 486 images | aalh_iit_parksnature_001

2021-03-16 | 502 images | aalh_iit_buildings_008

2021-03-16 | 531 images | aalh_iit_buildings_009

2021-03-19 | 593 images | aalh_iit_buildings_010

2021-03-22 | 195 images | aalh_iit_celebrations_001

2021-03-23 | 37 images | aalh_iit_churches_002

2021-03-23 | 138 images | aalh_iit_parksnature_002
