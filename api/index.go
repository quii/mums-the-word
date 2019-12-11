package mumstheword

import (
	"archive/zip"
	"bytes"
	"encoding/json"
	"encoding/xml"
	"fmt"
	"log"
	"net/http"
	"net/http/httputil"
)

type WordStructure struct {
	Title string
	Body  string
}

func Handler(w http.ResponseWriter, r *http.Request) {
	dump, _ := httputil.DumpRequest(r, true)
	fmt.Println(string(dump))
	file, header, err := r.FormFile("doc")

	var title bytes.Buffer
	var body bytes.Buffer

	if err != nil {
		http.Error(w, err.Error(), http.StatusBadRequest)
		return
	}

	reader, err := zip.NewReader(file, header.Size)

	if err != nil {
		http.Error(w, err.Error(), http.StatusBadRequest)
		return
	}

	w.Header().Set("Access-Control-Allow-Origin", "*")

	for _, f := range reader.File {
		if f.Name == "word/document.xml" {
			file, err := f.Open()

			var stuff Document
			err = xml.NewDecoder(file).Decode(&stuff)

			if err != nil {
				log.Fatal(err)
			}

			for _, thing := range stuff.Body.P[0].R {
				fmt.Fprintf(&title, thing.T.Text)
			}

			for _, thing := range stuff.Body.P[1:] {
				for _, r := range thing.R {
					fmt.Fprintf(&body, r.T.Text)
				}
			}
		}
	}

	w.Header().Add("content-type", "application/json")

	json.NewEncoder(w).Encode(WordStructure{
		Title: title.String(),
		Body:  body.String(),
	})
}

/*
	// Open a zip archive for reading.
	r, err := zip.OpenReader("test.docx")
	if err != nil {
		log.Fatal(err)
	}
	defer r.Close()

	// Iterate through the files in the archive,
	// printing some of their contents.
	for _, f := range r.File {
		if f.Name == "word/document.xml" {
			file, err := f.Open()
			var stuff Document
			err = xml.NewDecoder(file).Decode(&stuff)

			if err != nil {
				log.Fatal(err)
			}

			for _, thing := range stuff.Body.P[0].R {
				fmt.Println("xx", thing.T.Text)
			}
		}
	}
*/

type Document struct {
	XMLName   xml.Name `xml:"document"`
	Text      string   `xml:",chardata"`
	Wpc       string   `xml:"wpc,attr"`
	Cx        string   `xml:"cx,attr"`
	Cx1       string   `xml:"cx1,attr"`
	Cx2       string   `xml:"cx2,attr"`
	Cx3       string   `xml:"cx3,attr"`
	Cx4       string   `xml:"cx4,attr"`
	Cx5       string   `xml:"cx5,attr"`
	Cx6       string   `xml:"cx6,attr"`
	Cx7       string   `xml:"cx7,attr"`
	Cx8       string   `xml:"cx8,attr"`
	Mc        string   `xml:"mc,attr"`
	Aink      string   `xml:"aink,attr"`
	Am3d      string   `xml:"am3d,attr"`
	O         string   `xml:"o,attr"`
	R         string   `xml:"r,attr"`
	M         string   `xml:"m,attr"`
	V         string   `xml:"v,attr"`
	Wp14      string   `xml:"wp14,attr"`
	Wp        string   `xml:"wp,attr"`
	W10       string   `xml:"w10,attr"`
	W         string   `xml:"w,attr"`
	W14       string   `xml:"w14,attr"`
	W15       string   `xml:"w15,attr"`
	W16cid    string   `xml:"w16cid,attr"`
	W16se     string   `xml:"w16se,attr"`
	Wpg       string   `xml:"wpg,attr"`
	Wpi       string   `xml:"wpi,attr"`
	Wne       string   `xml:"wne,attr"`
	Wps       string   `xml:"wps,attr"`
	Ignorable string   `xml:"Ignorable,attr"`
	Body      struct {
		Text string `xml:",chardata"`
		P    []struct {
			Text         string `xml:",chardata"`
			ParaId       string `xml:"paraId,attr"`
			TextId       string `xml:"textId,attr"`
			RsidR        string `xml:"rsidR,attr"`
			RsidRPr      string `xml:"rsidRPr,attr"`
			RsidRDefault string `xml:"rsidRDefault,attr"`
			RsidP        string `xml:"rsidP,attr"`
			PPr          struct {
				Text   string `xml:",chardata"`
				PStyle struct {
					Text string `xml:",chardata"`
					Val  string `xml:"val,attr"`
				} `xml:"pStyle"`
				Spacing struct {
					Text              string `xml:",chardata"`
					Before            string `xml:"before,attr"`
					BeforeAutospacing string `xml:"beforeAutospacing,attr"`
					After             string `xml:"after,attr"`
					AfterAutospacing  string `xml:"afterAutospacing,attr"`
					Line              string `xml:"line,attr"`
					LineRule          string `xml:"lineRule,attr"`
				} `xml:"spacing"`
				RPr struct {
					Text   string `xml:",chardata"`
					RFonts struct {
						Text     string `xml:",chardata"`
						Ascii    string `xml:"ascii,attr"`
						EastAsia string `xml:"eastAsia,attr"`
						HAnsi    string `xml:"hAnsi,attr"`
						Cs       string `xml:"cs,attr"`
					} `xml:"rFonts"`
					Sz struct {
						Text string `xml:",chardata"`
						Val  string `xml:"val,attr"`
					} `xml:"sz"`
					SzCs struct {
						Text string `xml:",chardata"`
						Val  string `xml:"val,attr"`
					} `xml:"szCs"`
					Lang struct {
						Text     string `xml:",chardata"`
						EastAsia string `xml:"eastAsia,attr"`
					} `xml:"lang"`
					I     string `xml:"i"`
					ICs   string `xml:"iCs"`
					Color struct {
						Text string `xml:",chardata"`
						Val  string `xml:"val,attr"`
					} `xml:"color"`
				} `xml:"rPr"`
				NumPr struct {
					Text string `xml:",chardata"`
					Ilvl struct {
						Text string `xml:",chardata"`
						Val  string `xml:"val,attr"`
					} `xml:"ilvl"`
					NumId struct {
						Text string `xml:",chardata"`
						Val  string `xml:"val,attr"`
					} `xml:"numId"`
				} `xml:"numPr"`
				TextAlignment struct {
					Text string `xml:",chardata"`
					Val  string `xml:"val,attr"`
				} `xml:"textAlignment"`
				Ind struct {
					Text string `xml:",chardata"`
					Left string `xml:"left,attr"`
				} `xml:"ind"`
			} `xml:"pPr"`
			BookmarkStart struct {
				Text string `xml:",chardata"`
				ID   string `xml:"id,attr"`
				Name string `xml:"name,attr"`
			} `xml:"bookmarkStart"`
			R []struct {
				Text    string `xml:",chardata"`
				RsidRPr string `xml:"rsidRPr,attr"`
				T       struct {
					Text  string `xml:",chardata"`
					Space string `xml:"space,attr"`
				} `xml:"t"`
				RPr struct {
					Text   string `xml:",chardata"`
					RFonts struct {
						Text     string `xml:",chardata"`
						Ascii    string `xml:"ascii,attr"`
						EastAsia string `xml:"eastAsia,attr"`
						HAnsi    string `xml:"hAnsi,attr"`
						Cs       string `xml:"cs,attr"`
					} `xml:"rFonts"`
					B   string `xml:"b"`
					BCs string `xml:"bCs"`
					Sz  struct {
						Text string `xml:",chardata"`
						Val  string `xml:"val,attr"`
					} `xml:"sz"`
					SzCs struct {
						Text string `xml:",chardata"`
						Val  string `xml:"val,attr"`
					} `xml:"szCs"`
					Lang struct {
						Text     string `xml:",chardata"`
						EastAsia string `xml:"eastAsia,attr"`
					} `xml:"lang"`
					NoProof string `xml:"noProof"`
					Color   struct {
						Text string `xml:",chardata"`
						Val  string `xml:"val,attr"`
					} `xml:"color"`
					Bdr struct {
						Text  string `xml:",chardata"`
						Val   string `xml:"val,attr"`
						Sz    string `xml:"sz,attr"`
						Space string `xml:"space,attr"`
						Color string `xml:"color,attr"`
						Frame string `xml:"frame,attr"`
					} `xml:"bdr"`
					I   string `xml:"i"`
					ICs string `xml:"iCs"`
					Shd struct {
						Text  string `xml:",chardata"`
						Val   string `xml:"val,attr"`
						Color string `xml:"color,attr"`
						Fill  string `xml:"fill,attr"`
					} `xml:"shd"`
				} `xml:"rPr"`
				Br struct {
					Text string `xml:",chardata"`
					Type string `xml:"type,attr"`
				} `xml:"br"`
				Drawing struct {
					Text   string `xml:",chardata"`
					Inline struct {
						Text     string `xml:",chardata"`
						DistT    string `xml:"distT,attr"`
						DistB    string `xml:"distB,attr"`
						DistL    string `xml:"distL,attr"`
						DistR    string `xml:"distR,attr"`
						AnchorId string `xml:"anchorId,attr"`
						EditId   string `xml:"editId,attr"`
						Extent   struct {
							Text string `xml:",chardata"`
							Cx   string `xml:"cx,attr"`
							Cy   string `xml:"cy,attr"`
						} `xml:"extent"`
						EffectExtent struct {
							Text string `xml:",chardata"`
							L    string `xml:"l,attr"`
							T    string `xml:"t,attr"`
							R    string `xml:"r,attr"`
							B    string `xml:"b,attr"`
						} `xml:"effectExtent"`
						DocPr struct {
							Text string `xml:",chardata"`
							ID   string `xml:"id,attr"`
							Name string `xml:"name,attr"`
						} `xml:"docPr"`
						CNvGraphicFramePr struct {
							Text              string `xml:",chardata"`
							GraphicFrameLocks struct {
								Text           string `xml:",chardata"`
								A              string `xml:"a,attr"`
								NoChangeAspect string `xml:"noChangeAspect,attr"`
							} `xml:"graphicFrameLocks"`
						} `xml:"cNvGraphicFramePr"`
						Graphic struct {
							Text        string `xml:",chardata"`
							A           string `xml:"a,attr"`
							GraphicData struct {
								Text string `xml:",chardata"`
								URI  string `xml:"uri,attr"`
								Pic  struct {
									Text    string `xml:",chardata"`
									Pic     string `xml:"pic,attr"`
									NvPicPr struct {
										Text  string `xml:",chardata"`
										CNvPr struct {
											Text string `xml:",chardata"`
											ID   string `xml:"id,attr"`
											Name string `xml:"name,attr"`
										} `xml:"cNvPr"`
										CNvPicPr struct {
											Text     string `xml:",chardata"`
											PicLocks struct {
												Text               string `xml:",chardata"`
												NoChangeAspect     string `xml:"noChangeAspect,attr"`
												NoChangeArrowheads string `xml:"noChangeArrowheads,attr"`
											} `xml:"picLocks"`
										} `xml:"cNvPicPr"`
									} `xml:"nvPicPr"`
									BlipFill struct {
										Text string `xml:",chardata"`
										Blip struct {
											Text   string `xml:",chardata"`
											Embed  string `xml:"embed,attr"`
											ExtLst struct {
												Text string `xml:",chardata"`
												Ext  struct {
													Text        string `xml:",chardata"`
													URI         string `xml:"uri,attr"`
													UseLocalDpi struct {
														Text string `xml:",chardata"`
														A14  string `xml:"a14,attr"`
														Val  string `xml:"val,attr"`
													} `xml:"useLocalDpi"`
												} `xml:"ext"`
											} `xml:"extLst"`
										} `xml:"blip"`
										SrcRect string `xml:"srcRect"`
										Stretch struct {
											Text     string `xml:",chardata"`
											FillRect string `xml:"fillRect"`
										} `xml:"stretch"`
									} `xml:"blipFill"`
									SpPr struct {
										Text   string `xml:",chardata"`
										BwMode string `xml:"bwMode,attr"`
										Xfrm   struct {
											Text string `xml:",chardata"`
											Off  struct {
												Text string `xml:",chardata"`
												X    string `xml:"x,attr"`
												Y    string `xml:"y,attr"`
											} `xml:"off"`
											Ext struct {
												Text string `xml:",chardata"`
												Cx   string `xml:"cx,attr"`
												Cy   string `xml:"cy,attr"`
											} `xml:"ext"`
										} `xml:"xfrm"`
										PrstGeom struct {
											Text  string `xml:",chardata"`
											Prst  string `xml:"prst,attr"`
											AvLst string `xml:"avLst"`
										} `xml:"prstGeom"`
										NoFill string `xml:"noFill"`
										Ln     struct {
											Text   string `xml:",chardata"`
											NoFill string `xml:"noFill"`
										} `xml:"ln"`
									} `xml:"spPr"`
								} `xml:"pic"`
							} `xml:"graphicData"`
						} `xml:"graphic"`
					} `xml:"inline"`
				} `xml:"drawing"`
				LastRenderedPageBreak string `xml:"lastRenderedPageBreak"`
			} `xml:"r"`
			ProofErr []struct {
				Text string `xml:",chardata"`
				Type string `xml:"type,attr"`
			} `xml:"proofErr"`
			Hyperlink []struct {
				Text     string `xml:",chardata"`
				ID       string `xml:"id,attr"`
				TgtFrame string `xml:"tgtFrame,attr"`
				History  string `xml:"history,attr"`
				R        []struct {
					Text    string `xml:",chardata"`
					RsidRPr string `xml:"rsidRPr,attr"`
					RPr     struct {
						Text   string `xml:",chardata"`
						RFonts struct {
							Text     string `xml:",chardata"`
							Ascii    string `xml:"ascii,attr"`
							EastAsia string `xml:"eastAsia,attr"`
							HAnsi    string `xml:"hAnsi,attr"`
							Cs       string `xml:"cs,attr"`
						} `xml:"rFonts"`
						Color struct {
							Text string `xml:",chardata"`
							Val  string `xml:"val,attr"`
						} `xml:"color"`
						Sz struct {
							Text string `xml:",chardata"`
							Val  string `xml:"val,attr"`
						} `xml:"sz"`
						SzCs struct {
							Text string `xml:",chardata"`
							Val  string `xml:"val,attr"`
						} `xml:"szCs"`
						U struct {
							Text string `xml:",chardata"`
							Val  string `xml:"val,attr"`
						} `xml:"u"`
						Lang struct {
							Text     string `xml:",chardata"`
							EastAsia string `xml:"eastAsia,attr"`
						} `xml:"lang"`
						B   string `xml:"b"`
						BCs string `xml:"bCs"`
						I   string `xml:"i"`
						ICs string `xml:"iCs"`
					} `xml:"rPr"`
					T struct {
						Text  string `xml:",chardata"`
						Space string `xml:"space,attr"`
					} `xml:"t"`
				} `xml:"r"`
				ProofErr []struct {
					Text string `xml:",chardata"`
					Type string `xml:"type,attr"`
				} `xml:"proofErr"`
			} `xml:"hyperlink"`
		} `xml:"p"`
		BookmarkEnd struct {
			Text string `xml:",chardata"`
			ID   string `xml:"id,attr"`
		} `xml:"bookmarkEnd"`
		SectPr struct {
			Text  string `xml:",chardata"`
			RsidR string `xml:"rsidR,attr"`
			PgSz  struct {
				Text string `xml:",chardata"`
				W    string `xml:"w,attr"`
				H    string `xml:"h,attr"`
			} `xml:"pgSz"`
			PgMar struct {
				Text   string `xml:",chardata"`
				Top    string `xml:"top,attr"`
				Right  string `xml:"right,attr"`
				Bottom string `xml:"bottom,attr"`
				Left   string `xml:"left,attr"`
				Header string `xml:"header,attr"`
				Footer string `xml:"footer,attr"`
				Gutter string `xml:"gutter,attr"`
			} `xml:"pgMar"`
			Cols struct {
				Text  string `xml:",chardata"`
				Space string `xml:"space,attr"`
			} `xml:"cols"`
			DocGrid struct {
				Text      string `xml:",chardata"`
				LinePitch string `xml:"linePitch,attr"`
			} `xml:"docGrid"`
		} `xml:"sectPr"`
	} `xml:"body"`
}
