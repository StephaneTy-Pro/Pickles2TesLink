package main

import (
    "encoding/csv"
    "encoding/json"
    "fmt"
    "io/ioutil"
    "os"
    "strings"
    "github.com/aswjh/excel"
    "github.com/cheggaaa/pb"
    "time"    
    "encoding/xml"
    "io"
    "log"
    "path/filepath"
    "unsafe"
    //"strconv"
)


type PicklesFeature struct {
    Name string `json:"Name"`
	Description string `json:"Description"`
	FeatureElements []PicklesFeatureElement
    WasSuccessful   bool `json:"WasSuccessful"`
}
type PicklesFeatureElement struct {
	Name string `json:"Name"`
	Description string `json:"Description"`
    Steps []PicklesStep
    Tags [] string `json:"Tags"`
    Result PicklesResult
}
type PicklesStep struct {
	Keyword string `json:"Keyword"`
	NativeKeyword string `json:"NativeKeyword"`
	Name string `json:"Name"`
    TableArgument PicklesTableArgument
}
type PicklesTableArgument struct {
    HeaderRow []string `json:"HeaderRow"`
    DataRows [][]string `json:"DataRows"`

}
type PicklesResult struct {
    WasExecuted     bool `json:"WasExecuted"`
    WasSuccessful   bool `json:"WasSuccessful"`
}
type Pickles struct {
	Features []struct {
	   RelativeFolder string `json:"RelativeFolder"`
	   Feature PicklesFeature
	   Result PicklesResult
    } `json:"Features"`
	Configuration struct {
	   GeneratedOn string `json:"GeneratedOn"`
	}
}

type TstLnk struct {
	Entity string `xml:"entity,attr"`  
}


type TestLinkXls struct {
   TsName string
   TsDetails string
   TcName string
   TcSummary string
   TcPreCond string
   TcImportant string
   TcStep string
   TcExpRes string
   TcDocId string
}


type TestSuite struct {
	XMLName   xml.Name `xml:"testsuite"`
	Name      string   `xml:"name,attr"`
	Details   CharData `xml:"details"`
	TestCase  []TestCase
	TestSuite []TestSuite `xml:omitempty`
	Comment   string      `xml:",comment"`
}

type CharData string

func (n CharData) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	return e.EncodeElement(struct {
		S string `xml:",innerxml"`
	}{
		S: "<![CDATA[" + string(n) + "]]>",
	}, start)
}




type TestCase struct {
	XMLName       xml.Name      `xml:"testcase"`
	Name          string        `xml:"name,attr"`
	ExternalId    CharData      `xml:"externalid"`
	Summary       CharData      `xml:"summary"`
	Preconditions CharData      `xml:"preconditions"`
	Importance    CharData      `xml:"importance"`
	CustomFields  []CustomField `xml:"custom_fields"`
	Steps         []Step        `xml:"steps>step"`     // permet de générer un seul STEPS mais plusieurs STEP a l'insertion d'un step si on mets step a la place, il genere autant de noueds steps qu'il y aura de noeuds step
	//ExpectedResults CharData     `xml:"expectedresults"`
	Keywords Keywords `xml:"keywords,omitempty"`
}

type CustomField struct {
	Name  CharData `xml:"custom_field>name"`
	Value CharData `xml:"custom_field>value"`
}

type Step struct {
	StepNumber      CharData `xml:"step_number"` // avant on avait step>step_number mais du fait de la correction a TestCase Steps on en a plus besoin
	Actions         CharData `xml:"actions"`
	ExpectedResults CharData `xml:"expectedresults"`
	ExecutionType   CharData `xml:"execution_type"`
}

type Keywords struct {
	Keyword []Keyword `xml:"keyword"`
}

type Keyword struct {
	Property string   `xml:"property,attr"`
	Name     string   `xml:"name,attr"`
	Notes    CharData `xml:"keyword>notes"`
}

type PtrPool struct {
	Name string
    Ptr unsafe.Pointer
}

var registry = map[string]PtrPool{}
// add a type to the registry
func registerType(t PtrPool) {
	name := t.Name
	registry[name] = t
}


func main() {
    cwd, _ := os.Getwd()
		if _, err := os.Stat(cwd + "/pickledFeatures.json"); os.IsNotExist(err) {
			fmt.Printf("Erreur irrecouvrable :  %v \n", err)
			os.Exit(1)
		}
    // reading data from JSON File
    data, err := ioutil.ReadFile(cwd + "/pickledFeatures.json")
    if err != nil {
        fmt.Println(err)
    }

    var d Pickles
    err = json.Unmarshal([]byte(data), &d)
    if err != nil {
        fmt.Println(err)
    }
    
    

    var lofTst []TestLinkXls

    //fmt.Printf("Configuration: %v\n", d.Configuration.GeneratedOn)
    fmt.Printf("Chargement du fichier JSON \n", )
    bar := pb.StartNew(len(d.Features))
    
	tests := &TestSuite{
		Name: "",
	}
    
    ExternalId := 1
  

    for k, v := range d.Features {
        var ts TestLinkXls

        path := strings.Replace(string(v.RelativeFolder), "\\", string(filepath.Separator), -1)
        dir, _ := filepath.Split(path)
        splittedDir := strings.Split(dir, string(filepath.Separator))
        deep := 0
        for _, i := range splittedDir {
            if len(i) > 0 {
                var ok bool
                key := fmt.Sprintf("%d_%s", deep, i)
                _, ok = registry[key]
                if ok == false {
                    Dir := TestSuite{ // a := &TestSuite{ ne fonctionne pas ici avec les pointer unsafe, quand je recupere la valeur je pointe sur du vide
                        Name: i,
                        Details: CharData("Repertoire"),
                        }
                    a := PtrPool{Name: key, Ptr: unsafe.Pointer(&Dir)}
                    registerType(a)
                }
                deep++

            }

        }
        
        ts.TsName = v.Feature.Name
        ts.TsDetails = v.Feature.Description
        lofTst = append(lofTst, ts )//.String())
        //appendXl(xl, ts, Line)
        fts := &TestSuite{
		  Name: ts.TsName,
          Details: CharData(ts.TsDetails),
	    }
        //fmt.Printf("ligne %d colonne 1 -> %s\n", Line, ts.TsName) 

        //fmt.Printf("Features[%d].Feature.Description: %v\n", k, v.Feature.Description)
        for l, w := range v.Feature.FeatureElements {
            var tc TestLinkXls  // j'ai besoin d'un nouvel element
            tc.TsName = "" //ts.TsName
            tc.TsDetails = "" //ts.TsDetails
            //fmt.Printf("FeatureElements[%d].Name: %v\n", l,w.Name)
            tc.TcName = w.Name
            //fmt.Printf("FeatureElements[%d].Description: %v\n", l, v.Feature.Description)
            tc.TcSummary= w.Description
            
            ftc := &TestCase{
                Name: w.Name,
                Summary: CharData(w.Description),
                ExternalId: CharData(fmt.Sprintf("%d",ExternalId)),
                Importance: CharData(fmt.Sprintf("%d",1)),
            } 
            /*
            fcc := &CustomField{       
                    Name: "<Custom 1>",
                    Value: "Document ID"}    
            ftc.CustomFields = append(ftc.CustomFields, *fcc )         
            */        
            
            
            given, when, then := false, false, false
            iStep := 0
            NumberOfStep :=0
            Preconditions := "" // ne peut pas etre recuperé dans object Testlink car il est remise à zero a chaque "Quand" pour avoir une ligne sur le tableau excel


            for m, x := range w.Steps { 

               
                
                switch strings.ToLower(strings.Trim(x.NativeKeyword," ")) {
                case "etant donné":
                    // Attention pas de Given multiple dans un step sinon  NumberOfStep++ devrait etre ici
                    //fmt.Printf("je lis Given (iStep %d) \n", iStep)  
                    given = true
                    when = false

                    if then == true {
                        CreateStep(&tc,NumberOfStep-1,ftc) // c'est le step précédent que je veux ajouter 
                        lofTst = append(lofTst, tc ) 
                    }
                                                            
                    then = false                    
                    tc.TcPreCond = fmt.Sprintf("%s\n%s%s",tc.TcPreCond,x.NativeKeyword,x.Name)
                    tc.TcPreCond = fmt.Sprintf("%s\n%s",tc.TcPreCond,CreateTable(x.TableArgument))
                    Preconditions = tc.TcPreCond
                    

/*
                    if NumberOfStep > 1 {
                         log.Fatalff("Erreur fatale : Pas de reset du step et TestLink ne sait pas gérer plusieurs préconditions (On traitait la ligne chemin %v feature %v(%s), element %v(%s) step %d(%s)\n",v.RelativeFolder,k,v.Feature.Name,l,w.Name,m,x.NativeKeyword+x.Name)
                        }*/
                case "quand":  
                    //fmt.Printf("je lis When (iStep %d) \n", iStep)              
                    given = false
                    when = true
                    NumberOfStep++

                    if then == true  {
                        /* la commande précédente etait un then, je dois ajouter la boucle */
                        CreateStep(&tc,NumberOfStep-1,ftc) // c'est le step précédent que je veux ajouter
                        lofTst = append(lofTst, tc )   
                        //fmt.Printf("Ajout au Quand : %d (On traitait la ligne feature %v(%s), element %v(%s) StepNumber %d - ref iteration %d(%s)\n",iStep,k,v.Feature.Name,l,w.Name,NumberOfStep,m,x.NativeKeyword+x.Name)
                        if tc.TcStep =="" {
                            log.Fatalf("Erreur fatale : Aucun quand d'enregistré (On traitait la ligne feature %v(%s), element %v(%s) step %d(%s)\n",k,v.Feature.Name,l,w.Name,m,x.NativeKeyword+x.Name)
                        }
                        
                                          
                    }
                    then = false
                    tc=TestLinkXls{}
                    tc.TcStep = fmt.Sprintf("%s<br>%s%s",tc.TcStep,x.NativeKeyword,x.Name)
                    tc.TcStep = fmt.Sprintf("%s<br>%s",tc.TcStep,CreateTable(x.TableArgument))
                    iStep++

                case "et":            
                    //fmt.Printf("je lis And (iStep %d) \n", iStep) 
                    if given == true {
                        tc.TcPreCond = fmt.Sprintf("%s<br>%s%s",tc.TcPreCond,x.NativeKeyword,x.Name)
                        tc.TcPreCond = fmt.Sprintf("%s\n%s",tc.TcPreCond,CreateTable(x.TableArgument))
                        Preconditions = tc.TcPreCond 
                    }else if when == true {
                        tc.TcStep = fmt.Sprintf("%s<br>%s%s",tc.TcStep,x.NativeKeyword,x.Name)
                        tc.TcStep = fmt.Sprintf("%s<br>%s",tc.TcStep,CreateTable(x.TableArgument))
                        iStep++
                    } else if then == true {
                        tc.TcExpRes = fmt.Sprintf("%s<br>%s%s",tc.TcExpRes,x.NativeKeyword,x.Name)
                        tc.TcExpRes = fmt.Sprintf("%s<br>%s",tc.TcExpRes,CreateTable(x.TableArgument))
                        if len(w.Steps)-1 == m {
                            //fmt.Printf("Ajout au Then (dernier step) : %d (On traitait la ligne feature %v(%s), element %v(%s) StepNumber %d - ref iteration %d(%s)\n",iStep,k,v.Feature.Name,l,w.Name,NumberOfStep,m,x.NativeKeyword+x.Name)
                            CreateStep(&tc,NumberOfStep,ftc)  
                                lofTst = append(lofTst, tc )                  
                        }
                        iStep++        
                    } else {
                        log.Fatalf("Erreur fatale : Quand et Alors sont tous les deux actifs ce qui est impossible (On traitait la ligne feature %v(%s), element %v(%s) StepNumber %d - ref iteration %d(%s)\n",k,v.Feature.Name,l,w.Name,NumberOfStep,m,x.NativeKeyword+x.Name)
                    }
                    

                case "alors":
                    //fmt.Printf("je lis Then (iStep %d) \n", iStep)      
                    if when == false {
                        log.Fatalf("Erreur fatale : quand n'a pas ete specifié avant alors (On traitait la ligne feature %v(%s), element %v(%s) step %d(%s)\n",k,v.Feature.Name,l,w.Name,m,x.NativeKeyword+x.Name)
                    }
                    given = false
                    when = false
                    then = true
                    tc.TcExpRes= fmt.Sprintf("%s<br>%s%s",tc.TcExpRes,x.NativeKeyword,x.Name)
                    tc.TcExpRes = fmt.Sprintf("%s<br>%s",tc.TcExpRes,CreateTable(x.TableArgument))
                    if len(w.Steps)-1 == m {
                        //fmt.Printf("Ajout au dernier Step : %d (On traitait la ligne feature %v(%s), element %v(%s) StepNumber %d - ref iteration %d(%s)\n",iStep,k,v.Feature.Name,l,w.Name,NumberOfStep,m,x.NativeKeyword+x.Name)
                        CreateStep(&tc,NumberOfStep,ftc)  
                        lofTst = append(lofTst, tc ) 
                    }
                    iStep++
                default:
                    log.Fatalf("Erreur fatale : mot clé : [%s] non evalué (On traitait la ligne feature %v(%s), element %v(%s) step %d(%s)\n",x.NativeKeyword, k,v.Feature.Name,l,w.Name,m,x.NativeKeyword+x.Name)
                    iStep++
                }
                

                
                /* debug arreter la boucle sur un id 
                if (ExternalId == 99 && then==true) {

                    log.Fatal(" Id 99 ..... Exit")
                }                
                */ 
                
                
                
            }
            
            
                
            ftc.Preconditions = CharData(Preconditions)
                      

            if tc.TcName == w.Name {
                // je n'ai pas du écrire le test
                //fmt.Printf("Ajout de securite au dernier Step : %d ( %v)\n",iStep,tc.TcName )
                lofTst = append(lofTst, tc )
                ftp := &Step {
                    StepNumber: CharData(fmt.Sprintf("%d",NumberOfStep)),
                    Actions         : CharData(tc.TcStep),
                    ExpectedResults         : CharData(tc.TcExpRes),
                    ExecutionType: CharData(fmt.Sprintf("%d",1)),
                } 
                ftc.Steps = append(ftc.Steps, *ftp )  
            }
        fts.TestCase = append(fts.TestCase, *ftc)
        ExternalId++     
        }
        bar.Increment()
        time.Sleep(time.Millisecond)     
    
        
        for k,_ := range splittedDir {
            if k == 0 {
                continue
            }
                //fmt.Printf("boucle %v,%v,%v\n",k,len(s),s[len(s)-1-k]) // Suggestion: do `last := len(s)-1` before the loop
            key := fmt.Sprintf("%d_%s", len(splittedDir)-1-k, splittedDir[len(splittedDir)-1-k])
            //fmt.Printf("key %s\n",key)
            //fmt.Printf("registry['%s'] %v \n",key, registry[key])
            TsPtr := (*TestSuite)(registry[key].Ptr)
            //fmt.Printf("Acces a Name de registry['%s']\n", TsPtr.Name)
            
            TsPtr.TestSuite = append(TsPtr.TestSuite, *fts)
            fts = TsPtr
        }
            
        // ce sera pour quand on va gerer l'arbre des tests tests.TestSuite[k].TestSuite = []TestSuite{*fts}     
        tests.TestSuite = append(tests.TestSuite, *fts)         
                 
    }
    bar.FinishPrint("Done !")
    //fmt.Printf("Features[0].Result.WasExecuted: %v\n", d.Features[0].Result.WasExecuted)

    //fmt.Printf("test : %v",lofTst)
    
    /*
    out, err := xml.MarshalIndent(tests, "", "   ")

	if err != nil {
		panic(err)
	}

	fmt.Println(string(out))
    */
    
    filename := "test.xml" 
    file, _ := os.Create(cwd + "/"+filename) 
    xmlWriter := io.Writer(file) 
    defer file.Close()
    /*
    fonctionne bien mais sans header
    enc := xml.NewEncoder(xmlWriter) 
    enc.Indent("  ", "    ") 
    xmlWriter.Write(xml.Header)
    if err := enc.Encode(tests); err != nil { fmt.Printf("error: %v\n", err) } 
*/
/*
    out, err := xml.MarshalIndent(tests, "", "   ")

	if err != nil {
		panic(err)
	}
    
    xmlstring:= []byte(xml.Header + string(out))
	xmlWriter.Write(xmlstring)
*/

    
    if xmlstring, err := xml.MarshalIndent(tests, "", "    "); err == nil {
        xmlstring = []byte(xml.Header + string(xmlstring))
        //fmt.Printf("%s\n",xmlstring)
        xmlWriter.Write(xmlstring)
    }
    
    
    f, errw := os.Create(cwd + "/test.csv")

    if errw != nil {
        fmt.Printf("error opening dest csv:", err)
    }

    defer f.Close()

    w := csv.NewWriter(f)
    w.UseCRLF = true;
    w.Comma = ';'

    bar = pb.StartNew(len(lofTst))
    for _, r := range lofTst {
       // if err := w.Write(record.(string)); err != nil {
        //    fmt.Printf("error writing record to csv:", err)
        //fmt.Printf("%s;%s \n", r.TsName, r.TsDetails,r.TcName,r.TcSummary,r.TcPreCond,r.TcImportant,r.TcStep,r.TcExpRes)
        if err := w.Write([]string{r.TsName, r.TsDetails,r.TcName,r.TcSummary,r.TcPreCond,fmt.Sprintf("%d",r.TcImportant),r.TcStep,r.TcExpRes}); err != nil {
            fmt.Printf("error writing record to csv:", err)
        }
        bar.Increment()
        time.Sleep(time.Millisecond)           
        //}
    }

    w.Flush()
    
    bar.FinishPrint("Done !")

    fmt.Printf("Saved to %S\n", f.Name())
    
    os.Exit(1)
    
    fmt.Printf("Export dans fichier xls testlink\n")

    xl, _ := excel.Open(cwd+"\\02-ImportTestCasesIntoTestLink.xls", excel.Option{"Visible": false, "DisplayAlerts": false})
    
    defer xl.Quit()    
    
    sheet, err := xl.Sheet("TestSuiteXMLGeneration") // utilise une copie
    
    Line:=3
    bar = pb.StartNew(len(lofTst))
    for _, r := range lofTst {
            sheet.Cells(Line, 1, r.TsName)
            sheet.Cells(Line, 2, r.TsDetails)
            sheet.Cells(Line, 4, r.TcName)
            sheet.Cells(Line, 5, r.TcSummary)
            sheet.Cells(Line, 6, r.TcPreCond)
            sheet.Cells(Line, 8, r.TcImportant)
            sheet.Cells(Line, 11, r.TcStep)
            sheet.Cells(Line, 12, r.TcExpRes)
            sheet.Cells(Line, 15, r.TcDocId)
            Line++
            bar.Increment()
            time.Sleep(time.Millisecond)              
            
    }
   xl.SaveAs(cwd+"\\generated-testlink.xls")
   
   bar.FinishPrint("Done !")

}

func CreateStep(tc *TestLinkXls, NumberOfStep int, ftc *TestCase){
    ftp := &Step {
        StepNumber: CharData(fmt.Sprintf("%d",NumberOfStep)),
        Actions         : CharData(tc.TcStep),
        ExpectedResults         : CharData(tc.TcExpRes),
        ExecutionType: CharData(fmt.Sprintf("%d",1)),
        } 
    ftc.Steps = append(ftc.Steps, *ftp )
    //fmt.Printf("ftp %v\n\n", ftp)
    //fmt.Printf("ftc %v\n\n", ftc)    
}

func CreateTable(tbl PicklesTableArgument) string {
    Html:=""
    if len(tbl.HeaderRow)> 0 {
        Html+="<table>" 
        for n, y := range tbl.DataRows {
            if n == 1 && y[0] == strings.Repeat("-", strings.Count(y[0],"")-1) {continue}
            if n == len(tbl.DataRows)-1 && y[0] == strings.Repeat("-", strings.Count(y[0],"")-1) {continue}  
            //fmt.Printf("ligne %d grp[%s] chaine lue [%s] taille:[%d] chaine de tiret [%s]\n",n,y,y[0],strings.Count(y[0],""),strings.Repeat("-", strings.Count(y[0],"")))    
            Html+="<tr>" 
            for _, z := range y {
                if n == 0 {
                    Html += fmt.Sprintf("<th>%s</th>",z)
                    } else {
                    Html += fmt.Sprintf("<td>%s</td>",z)
                    }
                }
            Html+="</tr>"
            //fmt.Printf("TableArgument.DataRows[%d]: %v\n", n,y)
        }
        Html+="</table>" 
    // debug log.Fatalf("Stop %s",Html)
    
    }
    return Html
}

/* permet d'avoir une valeur par défaut*/
// dans ce cas la faire
// d := NewTestLinkXls()
// lofTst = append(lofTst ,*d)
func NewTestLinkXls() *TestLinkXls {
    return &TestLinkXls{TsName:"",TsDetails:""}
}


func appendXl(xl *excel.MSO, ts TestLinkXls, Line int) {
    //fmt.Printf("%v\n", xl)
    //fmt.Printf("%v\n", &xl)
    //fmt.Printf("%v\n", *xl) 
    // ex := *xl
    //sheet, err := ex.Sheet("TestSuiteXMLGeneration")

        
    //sheet, err := (*xl).Sheet("TestSuiteXMLGeneration")
    sheet, err := xl.Sheet("TestSuiteXMLGeneration") // utilise une copie
    if err != nil {
        fmt.Println("Error : %s",err)
    }    
    //fmt.Printf("ligne %d colonne 1 -> %s\n", Line, ts.TsName) 
    sheet.Cells(Line, 1, ts.TsName)
    sheet.Cells(Line, 2, ts.TsDetails)
    sheet.Cells(Line, 4, ts.TcName)
}