# Luminescence Dose and Age Calculator (LDAC v1.2) <img width=100px src="https://github.com/Peng-Liang/LDAC/blob/master/ICON/Small-Logo.png" align="right" />
[![License](https://img.shields.io/badge/license-MIT-brightgreen.svg)](https://github.com/Peng-Liang/LDAC/blob/master/LICENSE) ![Star](https://img.shields.io/github/stars/Peng-Liang/LDAC.svg) ![commit](https://img.shields.io/github/commits-since/Peng-Liang/LDAC/v1.0.svg) [![downloads](https://img.shields.io/github/downloads/Peng-Liang/LDAC/total?color=light&label=Downloads)](https://github.com/Peng-Liang/LDAC/releases) [![language](https://img.shields.io/badge/Language-VBA-orange.svg)](https://docs.microsoft.com/en-us/office/vba/api/overview/excel) ![build](https://img.shields.io/badge/build-passing-brightgreen.svg) [![version](https://img.shields.io/badge/version-1.2-blue)](https://github.com/Peng-Liang/LDAC/releases)
## 1. Introduction
* The **Luminescence Dose and Age Calculator (LDAC)** is a *Microsoft Excel Visual Basic for Application (VBA)*-based package which can be used to assemble OSL age information and associated calculations. This platform applies statistical models to determine equivalent dose (De) values and render corresponding OSL age estimates. This software is fully applicable for De measurements by single grain and aliquot regeneration (SAR) and thermal transfer OSL (TT-OSL) protocols. It could also be used to calculate the dose rate and final buried age for geology/archaeology samples.

* The most RECENT version (**LDAC v1.2**) has been released on *Mar 28, 2023*. 
* ***An issue with the 'Recyc. Ratio' input error due to cell protection has been fixed on Mar 29, 2023***

## 2. Citation
* Liang, P., Forman, S.L., 2019. [LDAC: An Excel-based program for luminescence equivalent dose and burial age calculations](http://ancienttl.org/ATL_37-2_2019/ATL_37-2_Liang_p21-40.pdf). *Ancient TL* 37 (2), 21-40. 

* Download: [*[Full Text](http://ancienttl.org/ATL_37-2_2019/ATL_37-2_Liang_p21-40.pdf)*].     Citation: [*[BibTex](http://ancienttl.org/ATL_37-2_2019/ATL_37-2_Liang_citation.bib)*]   [*[RIS](https://github.com/Peng-Liang/LDAC/blob/master/ICON/Liang_AncientTL.RIS)*]

<a href="http://ancienttl.org/ATL_37-2_2019/ATL_37-2_Liang_p21-40.pdf" target="_blank"><img src="https://github.com/Peng-Liang/LDAC/blob/master/ICON/Picture1.png" alt="LDAC_Ancient TL" width="800" /></a>

## 3. LDAC requirements
* LDAC requires *Microsoft Excel 2010* or higher version (e.g., 2013, 2016, 2019) for *Windows* computers. *[Microsoft Excel 2019](https://products.office.com/en-US/get-started-with-office-2019?&OCID=AID2000136_SEM_iNi8NhPm&MarinID=siNi8NhPm%7C340667806722%7Cmicrosoft%20office%202019%7Ce%7Cc%7C%7C54569958854%7Caud-473968998473:kwd-331146748204&lnkd=Google_O365SMB_NI&gclid=Cj0KCQjwvdXpBRCoARIsAMJSKqLLubP-daYYm88zMR_H2RSsXydSHLheCSbXj7UGBKynT_lqAtzqqlQaAuJ-EALw_wcB)* is highly recommended.

* *Macintosh Excel* can be used to preview the data, but the *Macros* cannot be run. A Windows-enabling program (e.g., *Fusion, Parallels*) is to run **LDAC**.

## 4. Download the LDAC ![size](https://img.shields.io/badge/Software%20size-6.89M-blue.svg)
* The **LDAC** is continuously being developed and improved. The most recent (*Mar 28, 2023*) distribution of LDAC can be downloaded [here (*![LDAC software (v1.2)](https://img.shields.io/badge/LDAC%20software-v1.2-brightgreen.svg)*)](https://github.com/Peng-Liang/LDAC/releases). 

* **Note: Extract the downloaded zip file** "*LDAC.software.v1.2.zip*" **and the** “*LDAC (v1.2).xlsm*” **will be found**.

* --------New features (LDAC v1.2)-----------
* (1) Calculate dose rate in batches
* (2) Add an option for inputting radionuclides in Bq/kg.

## 5. Getting started

* **Note: The protection password in LDAC is “;”, which is used to protect the worksheet from unintentional modifications.**

* Make sure the downloaded workbook’s name is “*LDAC (v1.2).xlsm*”. if not, rename it.

* Open the workbook just downloaded from the internet. A warning message will show “*PROTECTED VIEW Be careful-files from the internet can contain viruses. Unless you need to edit, it’s safer to stay in Protect-ed view*”. Click “**Enable Editing**” to use this program. 

* On first running, the program **LDAC** might appear the following message “*SECURITY 
WARNING Some active content has been disabled. Click for more details*.” This is a warning message for using *Macros* and command buttons (ActiveX controls) of the Excel workbook. Click “**Enable Content**”. 

* If this warning message cannot be displayed and any button on the worksheet does not respond, check the *macro settings* in the Trust Center (“*Excel>File>Options>Trust Center>Trust Center Settings>Macro settings*”). [Enabling or disabling Macros in Excel refer to the support document from the Microsoft website](https://support.office.com/en-us/article/enable-or-disable-macros-in-office-files-12b036fd-d140-4e74-b45e-16fed1a7e5c6).

* A [training video](https://youtu.be/Of_feY1UeqU) can be viewed on Youtube.
<a href="https://youtu.be/Of_feY1UeqU" target="_blank"><img src="https://github.com/Peng-Liang/LDAC/blob/master/ICON/Video_Still.png" alt="LDACTrain" width="600" height="337" border="30" /></a>

## 6. Feedback
* Although we have tried this program in lots of computers with different language version of *Windows* and *Microsoft Excels*, we believe that users may still encounter some unknown errors and bugs. 

* Any bug-reports, suggestions, and even requirements for further developing the LDAC are warmly welcome. Please contact Peng Liang (PLiang@zju.edu.cn; LiangPeng2012@live.cn). I will get back to you as soon as possible.

## 7. Acknowledgments
* This work was supported by the China Scholarship Council (awarded to P.L.), the National Natural Science Foundation of China (#41430532), the State Scientific Survey Project of China (#2017FY101001), USA National Science Foundation Award #GSS-1660230 (SLF), National Geographic Society Award #9990-16 (SLF), and the Geoluminescence Dating Research Laboratory at Baylor University, USA. Sincere thanks are extended to Liliana Marín for helpful discussions and suggestions.

## 8. Featured publications using LDAC
* Li, G., Zhang, H., Liu, X., Yang, H., Wang, X., Zhang, X., ... & Xia, D. (2020). [Paleoclimatic changes and modulation of East Asian summer monsoon by high-latitude forcing over the last 130,000 years as revealed by independently dated loess-paleosol sequences on the NE Tibetan Plateau](https://doi.org/10.1016/j.quascirev.2020.106283). Quaternary Science Reviews, 237, 106283.
* Li, G., Wang, Z., Zhao, W., Jin, M., Wang, X., Tao, S., ... & Madsen, D. (2020). [Quantitative precipitation reconstructions from Chagan Nur revealed lag response of East Asian summer monsoon precipitation to summer insolation during the Holocene in arid northern China](https://doi.org/10.1016/j.quascirev.2020.106365). Quaternary Science Reviews, 239, 106365.
* Yang, H., Li, G., Huang, X., Wang, X., Zhang, Y., Jonell, T. N., ... & Deng, Y. (2020). [Loess depositional dynamics and paleoclimatic changes in the Yili Basin, Central Asia, over the past 250 ka](https://doi.org/10.1016/j.catena.2020.104881). Catena, 195, 104881.
* Yang, S., Liu, N., Li, D., Cheng, T., Liu, W., Li, S., ... & Luo, Y. (2021). [Quartz OSL chronology of the loess deposits in the Western Qinling Mountains, China, and their palaeoenvironmental implications since the Last Glacial period](https://doi.org/10.1111/bor.12473). Boreas, 50(1), 294-307.
* Liu, L., Yang, S., Cheng, T., Liu, X., Luo, Y., Liu, N., ... & Liu, W. (2021). [Chronology and dust mass accumulation history of the Wenchuan loess on eastern Tibetan Plateau since the last glacial](https://doi.org/10.1016/j.aeolia.2021.100748). Aeolian Research, 53, 100748.
* Ramírez-Herrera, M. T., Gaidzik, K., & Forman, S. L. (2021). [Spatial Variations of Tectonic Uplift-Subducting Plate Effects on the Guerrero Forearc, Mexico](https://doi.org/10.3389/feart.2020.573081). Frontiers in Earth Science, 590.
* Bollinger, L., Klinger, Y., Forman, S. L., Chimed, O., Bayasgalan, A., Munkhuu, U., ... & Sodnomsambuu, D. (2021). [25,000 Years long seismic cycle in a slow deforming continental region of Mongolia](https://doi.org/10.1038/s41598-021-97167-w). Scientific reports, 11(1), 17855.
* Abbas, W., Zhang, J., Tsukamoto, S., Ali, S., Frechen, M., & Reicherter, K. (2022). [Pleistocene-Holocene deformation and seismic history of the Kalabagh Fault in Pakistan using OSL and post-IR IRSL dating](https://doi.org/10.1016/j.quaint.2022.01.007). Quaternary International.651, 42-61.
* Zhang S, Zhao H, Sheng Y, Zhang J, Zhang J, Sun A, Wang L, Huang L, Hou J, Chen F. 2022. Mega-lakes in the northwestern Tibetan Plateau formed by melting glaciers during the last deglacial. Quaternary Science Reviews, 285: 107528.
* Zhang, S., Zhao, H., Sheng, Y., Chen, S., Li, G., Chen, F., 2022. Late Quaternary lake level record of Orog Nuur, southern Mongolia, revealed by optical dating of paleo-shorelines. Quaternary Geochronology, 72, 101370.
* Han, Y.-S., Zhang, J.-F., Liu, G.-N., Cui, Z.-J., 2022. Optical Dating of Quartz Grains From the Minjiang Fluvial Terraces in the Sonpan Area on the Eastern Margin of the Tibetan Plateau. Frontiers in Earth Science, 10, 939539.
* Forman S L, Tew-Todd V, Mayhack C, Marín L, Wiest L A, Money G., 2022. Late Quaternary aeolian environments, luminescence chronology and climate change for the Monahans dune field, Winkler County, West Texas, USA. Aeolian Research, 58: 100828.
* Özpolat, E., Yıldırım, C., Görüm, T., Gosse, J.C., Şahiner, E., Sarıkaya, M.A., Owen, L.A., 2022.Three-dimensional control of alluvial fans by rock uplift in an extensional regime: Aydın Range, Aegean extensional province. Scientific Reports, 12, 15306.
* Davis, L.G., Madsen, D.B., Sisson, D.A., Valdivia-Becerra, L., Higham, T., Stueber, D., Bean, D.W., Nyers, A.J., Carroll, A., Ryder, C., 2022. Dating of a large tool assemblage at the Cooper’s Ferry site (Idaho, USA) to~ 15,785 cal yr BP extends the age of stemmed points in the Americas. Science Advances, 8, eade1248.
* Mahan, S.A., Rittenour, T.M., Nelson, M.S., Ataee, N., Brown, N., DeWitt, R., Durcan, J., Evans, M., Feathers, J., Frouin, M., Guérin, G., Heydari, M., Huot, S., Jain, M., Keen-Zebert, A., Li, B., López, G.I., Neudorf, C., Porat, N., Rodrigues, K., Sawakuchi, A.O., Spencer, J.Q.G., Thomsen, K., 2022. Guide for interpreting and reporting luminescence dating results. GSA Bulletin.
* Mescolotti, P.C., Giannini, P.C.F., Pupim, F.d.N., Sawakuchi, A.O., Ladeira, F.S.B., Assine, M.L., 2023. The largest Quaternary inland eolian system in Brazil: Eolian landforms and activation/stabilization phases of the Xique-Xique dune field. Geomorphology, 420, 108516.

*(updated 3/28/2023 by P.L.)*
