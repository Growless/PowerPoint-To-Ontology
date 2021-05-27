package appPackage;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ooxml.POIXMLProperties.CoreProperties;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFAutoShape;
import org.apache.poi.xslf.usermodel.XSLFBackground;
import org.apache.poi.xslf.usermodel.XSLFHyperlink;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.semanticweb.owlapi.apibinding.OWLManager;
import org.semanticweb.owlapi.model.IRI;
import org.semanticweb.owlapi.model.OWLAnnotation;
import org.semanticweb.owlapi.model.OWLAnnotationAssertionAxiom;
import org.semanticweb.owlapi.model.OWLAnnotationProperty;
import org.semanticweb.owlapi.model.OWLClass;
import org.semanticweb.owlapi.model.OWLClassAssertionAxiom;
import org.semanticweb.owlapi.model.OWLDataFactory;
import org.semanticweb.owlapi.model.OWLDataProperty;
import org.semanticweb.owlapi.model.OWLDataPropertyAssertionAxiom;
import org.semanticweb.owlapi.model.OWLNamedIndividual;
import org.semanticweb.owlapi.model.OWLObjectProperty;
import org.semanticweb.owlapi.model.OWLObjectPropertyAssertionAxiom;
import org.semanticweb.owlapi.model.OWLOntology;
import org.semanticweb.owlapi.model.OWLOntologyCreationException;
import org.semanticweb.owlapi.model.OWLOntologyManager;
import org.semanticweb.owlapi.model.OWLOntologyStorageException;
import org.semanticweb.owlapi.model.PrefixManager;
import org.semanticweb.owlapi.reasoner.OWLReasoner;
import org.semanticweb.owlapi.reasoner.OWLReasonerFactory;
import org.semanticweb.owlapi.util.DefaultPrefixManager;

import uk.ac.manchester.cs.jfact.JFactFactory;



public class PptToOntology {

	private static final String[] type = {"XSLFAutoShape", "XSLFBackground", "XSLFChart", "XSLFComment",
			"XSLFCommentAuthors", "XSLFConnectorShape","XSLFDrawing","XSLFFontInfo","XSLFFreeformShape",
			"XSLFGraphicFrame", "XSLFGroupShape", "XSLFHyperlink", "XSLFMetroShape",
			"XSLFPictureShape","XSLFShadow","XSLFSimpleShape","XSLFTable","XSLFTextBox","XSLFTextShape",
			"XSLFTheme"};

	private static OWLOntology ontology;
	private static OWLReasonerFactory reasonerFactory = null;
	private static OWLReasoner reasoner = null;
	private static OWLOntologyManager ontManager;
	private static OWLDataFactory dataFactory = null;  //para las clases e instancias
	private static PrefixManager pm = null; //pm de ontología
	private static PrefixManager pm2 = new DefaultPrefixManager("http://www.essepuntato.it/2008/12/pattern"); //pm auxiliar para imports
	private static PrefixManager pm3 = new DefaultPrefixManager("http://purl.org/dc/elements/1.1/");
	private static PrefixManager pm4 = new DefaultPrefixManager("http://purl.org/spar/fabio");
	private static PrefixManager pm5 = new DefaultPrefixManager("http://purl.org/co");
	
	//nº de slides
	private static int nSlides = 1;
	//nº de objetos tipo texto
	private static int nText = 0;
    //nº de imagenes
	private static int nImg = 0;
    //nº de tablas
	private static int nTab = 0;
    //nº de tipos de autoshape
	private static int nAutoType = 0;
    //nº de hipervinculos
	private static int nLink = 0;
	//nº de parrafos
	private static int nPar = 0;
	//nº de listas
	private static int nList = 0;
	//nº de elementos de listas 
	private static int nListElems = 0;
	//nº de objetos tipo shape
	private static int nShape = 0;
	
	public static void main(String[] args) {
		
		String fileName;
		String ontology = null;
		String owlLocation = null;
		
		if (args.length > 0 && args.length <= 3) {
			
			if(args.length == 0) {
				
				fileName = args[0];
			}else if(args.length == 2){
				
				fileName = args[0];
				ontology = args[1];
			}else {
				fileName = args[0];
				ontology = args[1];
				owlLocation = args[2];
			}
			
		} else {
			
			System.out.println("No se ha especificado archivo, "
					+ "ontologia o lugar de destino del .owl");
			return;
		}
			
		FileInputStream inputStream;
		
		try {
			inputStream = new FileInputStream(fileName);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
			return;
		}
		
		XMLSlideShow ppt;
		
		try {
			ppt = new XMLSlideShow(inputStream);
		} catch (IOException e) {
			e.printStackTrace();
			return;
		}

		readSlide(ppt,ontology,owlLocation);
	}

	
	
	
	
	//Method that iterates between slides reading the shapes contained in them
	public static void readSlide(XMLSlideShow ppt, String ont, String owlLocation) {
		
        CoreProperties props = ppt.getProperties().getCoreProperties();
        String title = props.getTitle();
        System.out.println("Title: " + title);
        
        final String ONT_IRI = ont;
        ontManager = OWLManager.createOWLOntologyManager();
		IRI iri = IRI.create(ONT_IRI);
		try {
			ontology = ontManager.loadOntologyFromOntologyDocument(iri);
		} catch (OWLOntologyCreationException e) {
			e.printStackTrace();
		}
		
		reasonerFactory = new JFactFactory();
		reasoner = reasonerFactory.createReasoner(ontology);
		dataFactory = ontManager.getOWLDataFactory();
		pm = new DefaultPrefixManager(ONT_IRI);

        
		//////////////////Añadir presentación a ont//////////////////////
		OWLClass pptOnOnt = dataFactory.getOWLClass(":/Presentation",pm4);
		OWLNamedIndividual pptN = dataFactory.getOWLNamedIndividual(":/#presentation", pm4);
		OWLClassAssertionAxiom classAssertion1 = dataFactory.getOWLClassAssertionAxiom(pptOnOnt, pptN);
		ontManager.addAxiom(ontology, classAssertion1);
		
		
		////////////////////////////////////////////////////////////////
		
        //Bucle que recorre slides
        for (XSLFSlide slide: ppt.getSlides()) {
        	       	
        	List<XSLFShape> shapes = slide.getShapes();
        	int n = 0;

        	//////////////////Añadir las slides a la ontologia/////////////////
    		OWLClass slideOnOnt = dataFactory.getOWLClass(":/Part",pm);
   
    		//Se crea la class assertion para indicar que slide es una instancia de part
    		OWLNamedIndividual slideN = dataFactory.getOWLNamedIndividual(":/#slide" + nSlides, pm);
    		OWLClassAssertionAxiom classAssertion = dataFactory.getOWLClassAssertionAxiom(slideOnOnt, slideN);

    		//Se añade la class assertion a la ontologia
    		ontManager.addAxiom(ontology, classAssertion);
    		
    		//Añadimos la slide como "isContainedBy" de su ppt
      		OWLObjectProperty isContainedBy = dataFactory.getOWLObjectProperty("#isContainedBy",pm2);
      		 		
      		OWLObjectPropertyAssertionAxiom assertion =
      				dataFactory.getOWLObjectPropertyAssertionAxiom(isContainedBy, slideN, pptN);
      		
      		ontManager.addAxiom(ontology, assertion);
    		/////////////////////////////////////////////////////////////////
    		
    		/////////////////Añadir a la ontologia la background de cada slide/////////////////////
        	//background no es una shape
        	background(slide, slideN);
        	
        	for (XSLFShape shape: shapes) {
        		
        		//Si lo que llega se encuentra en el array de tipos admitidos
        		if ( contains(type,nombreTipo(shape.getClass().getName().toString())) ) {        			 
        			 n++;        			 
        			 //El indice de este switch coincide con el array "type"
	        		 switch (find(type,nombreTipo(shape.getClass().getName().toString()))) {
	        		
	        		 	case 0: //XSLFAutoShape
	        		 		autoshape(shape, slideN);
	        		 		break;
	        		 	case 1: //XSLFBackground
	        		 		//background(shape, slideN);
	        		 		break;
	        		 	case 13: //XSLFPictureShape
	        		 		nImg++;
	        		 		picture(shape, slideN);
	        		 		break;
	        		 	case 16: //XSLFTable
	        		 		nTab++;
	        		 		table(shape, slideN);
	        		 		break;
	        		 	case 17: //XSFLTextBox
	        		 		text(shape, slideN);
	        		 		break;
	        		 	case 18: //XSFLTextShape
	        		 		text(shape, slideN);
	        		 		break;    		 		
	        		 }
        		 }      		
        	}
        	
        	nSlides++;

        }
        
        System.out.println("Iniciando generación de .owl");
        
        //Creamos copia local de la ontología para visualizar en protege
        File file = new File(owlLocation + "/local.owl");
        try {
			ontManager.saveOntology(ontology, IRI.create(file.toURI()));
		} catch (OWLOntologyStorageException e) {
			
			e.printStackTrace();
		}
        
        System.out.println("Generado .owl");
	}
	
	
	
	
//--------------------------------------------------------------------------------------
	
	private static void background(XSLFSlide slide, OWLNamedIndividual parent) {
		
		XSLFBackground backShape = slide.getBackground();
		
		//Añadimos la backShape a la ontologia como parte de "Background" (no existe)
  		OWLClass backOnOnt = dataFactory.getOWLClass(":/Background",pm);
  		   
  		//Se crea la class assertion para indicar que back es una instancia de Background
  		OWLNamedIndividual bg = dataFactory.getOWLNamedIndividual(":/#background" + nSlides, pm);
  		OWLClassAssertionAxiom classAssertion = dataFactory.getOWLClassAssertionAxiom(backOnOnt, bg );

  		//Se añade la class assertion a la ontologia
  		ontManager.addAxiom(ontology, classAssertion);

  		//Añadimos el individual como "isContainedBy" de su slide
  		OWLObjectProperty isContainedBy = dataFactory.getOWLObjectProperty("#isContainedBy",pm2);
  		
  		OWLAnnotationProperty property = dataFactory.getOWLAnnotationProperty("/color", pm);
  		OWLAnnotation anno = dataFactory.getOWLAnnotation(property, dataFactory.getOWLLiteral(backShape.getFillColor().toString()));		
  		OWLAnnotationAssertionAxiom annoPropertyAssertion =
  								dataFactory.getOWLAnnotationAssertionAxiom(IRI.create("http://purl.org/spar/doco/#background" + nSlides), anno);
  		
  		OWLObjectPropertyAssertionAxiom assertion =
  				dataFactory.getOWLObjectPropertyAssertionAxiom(isContainedBy, bg, parent);
  		
  		ontManager.addAxiom(ontology, assertion);
  		ontManager.addAxiom(ontology, annoPropertyAssertion);

	}

	private static void table(XSLFShape shape, OWLNamedIndividual parent) {

	
		XSLFTable tabShape = (XSLFTable) shape;
		ArrayList<String> tableText = new ArrayList<String>();
		
		//Añadimos la tabShape a la ontologia como parte de "Table"
  		OWLClass tableOnOnt = dataFactory.getOWLClass(":/Table",pm);
  		   
  		//Se crea la class assertion para indicar que text es una instancia de table
  		OWLNamedIndividual table = dataFactory.getOWLNamedIndividual(":/#table" + nTab, pm);
  		OWLClassAssertionAxiom classAssertion = dataFactory.getOWLClassAssertionAxiom(tableOnOnt, table );

  		//Se añade la class assertion a la ontologia
  		ontManager.addAxiom(ontology, classAssertion);

  		//Añadimos el individual como "isContainedBy" de su slide
  		OWLObjectProperty isContainedBy = dataFactory.getOWLObjectProperty("#isContainedBy",pm2);
  		
  		OWLObjectPropertyAssertionAxiom assertion =
  				dataFactory.getOWLObjectPropertyAssertionAxiom(isContainedBy, table, parent);
  		
  		ontManager.addAxiom(ontology, assertion);


		
		//Recorrer la tabla para sacar contenido de las celdas
		for (int i=0; i < tabShape.getNumberOfRows(); i++) {
			
			for (int j = 0; j < tabShape.getNumberOfColumns(); j++) {
				tableText.add(tabShape.getCell(i, j).getText() + " "  +  "(" + i + "," + j + ")");
				
				if(!tabShape.getCell(i, j).getText().isBlank()) {
					nText++;
					text( tabShape.getCell(i, j),table);
				}	
			}
		}
		
		//Añadimos filas de la tabla como annotation
		 OWLAnnotationProperty property = dataFactory.getOWLAnnotationProperty("/tabRows", pm);
	  	 OWLAnnotation anno = dataFactory.getOWLAnnotation(property, dataFactory.getOWLLiteral(tabShape.getNumberOfRows()));		
	     OWLAnnotationAssertionAxiom annoPropertyAssertion =
	  		dataFactory.getOWLAnnotationAssertionAxiom(IRI.create("http://purl.org/spar/doco/#table" + nTab), anno);

	     ontManager.addAxiom(ontology, annoPropertyAssertion);
	     
		//Añadimos columnas de la tabla como annotation
	     property = dataFactory.getOWLAnnotationProperty("/tabColumns", pm);
	  	 anno = dataFactory.getOWLAnnotation(property, dataFactory.getOWLLiteral(tabShape.getNumberOfColumns()));		
	     annoPropertyAssertion =
	  		dataFactory.getOWLAnnotationAssertionAxiom(IRI.create("http://purl.org/spar/doco/#table" + nTab), anno);

	     ontManager.addAxiom(ontology, annoPropertyAssertion);
	     
		//Añadimos el texto que pudiese contener a la ontologia
		 property = dataFactory.getOWLAnnotationProperty("/tabContent", pm);
	  	 anno = dataFactory.getOWLAnnotation(property, dataFactory.getOWLLiteral(tableText.toString()));		
	     annoPropertyAssertion =
	  		dataFactory.getOWLAnnotationAssertionAxiom(IRI.create("http://purl.org/spar/doco/#table" + nTab), anno);

	     ontManager.addAxiom(ontology, annoPropertyAssertion);
		
	}

	private static void picture(XSLFShape shape, OWLNamedIndividual parent) {
		
		XSLFPictureShape picShape = null;
		XSLFAutoShape autoShape = null;
		OWLNamedIndividual pic = null;
		
		if(shape.getClass().getSimpleName().toString().equals("XSLFAutoShape")) {
			 autoShape = (XSLFAutoShape) shape;
			 //Aumentamos el numero de las shapes tipo "Forma"
			 nShape++;
			 pic = dataFactory.getOWLNamedIndividual(":/#shape" + nShape, pm);
			 
			 //movidas de forma
			 addShapeProperties(autoShape);
		}else {
			picShape = (XSLFPictureShape) shape;
			pic = dataFactory.getOWLNamedIndividual(":/#image" + nImg, pm);
			
			//movidas de imagen
			addImageProperties(picShape);
		  
		}
		//Añadimos la picShape a la ontologia como parte de "Figure"
  		OWLClass figure = dataFactory.getOWLClass(":/Figure",pm);
  		   
  		//Se crea la class assertion para indicar que text es una instancia de paragraph
  		
  		OWLClassAssertionAxiom classAssertion = dataFactory.getOWLClassAssertionAxiom(figure, pic);
  		
  		
  		//Se añade la class assertion a la ontologia
  		ontManager.addAxiom(ontology, classAssertion);
  		

  		//Añadimos el individual como "isContainedBy" de su slide
  		OWLObjectProperty isContainedBy = dataFactory.getOWLObjectProperty("#isContainedBy",pm2);
  		
  		OWLObjectPropertyAssertionAxiom assertion =
  				dataFactory.getOWLObjectPropertyAssertionAxiom(isContainedBy, pic, parent);
  		
  		ontManager.addAxiom(ontology, assertion);
		
	}

	private static void addImageProperties(XSLFPictureShape picShape) {
		
		//Añadimos el nombre de la imagen
		 OWLAnnotationProperty property = dataFactory.getOWLAnnotationProperty("title", pm3);
	  	 OWLAnnotation anno = dataFactory.getOWLAnnotation(property, dataFactory.getOWLLiteral(picShape.getPictureData().getFileName()));		
	     OWLAnnotationAssertionAxiom annoPropertyAssertion =
	  		dataFactory.getOWLAnnotationAssertionAxiom(IRI.create("http://purl.org/spar/doco/#image" + nSlides), anno);

	     ontManager.addAxiom(ontology, annoPropertyAssertion);
		
	     //Añadimos la extension de la imagen
	     property = dataFactory.getOWLAnnotationProperty("/extension", pm);
	     anno = dataFactory.getOWLAnnotation(property, dataFactory.getOWLLiteral(picShape.getPictureData().getType().extension));
	     annoPropertyAssertion =
	 	  		dataFactory.getOWLAnnotationAssertionAxiom(IRI.create("http://purl.org/spar/doco/#image" + nSlides), anno);
	     
	     ontManager.addAxiom(ontology, annoPropertyAssertion);
	     
	     //Añadimos el tamaño de la imagen
	     property = dataFactory.getOWLAnnotationProperty("/sizeInPixels", pm);
	     anno = dataFactory.getOWLAnnotation(property, dataFactory.getOWLLiteral(picShape.getPictureData().getImageDimensionInPixels().toString()));
	     annoPropertyAssertion =
	 	  		dataFactory.getOWLAnnotationAssertionAxiom(IRI.create("http://purl.org/spar/doco/#image" + nSlides), anno);
	     
	     ontManager.addAxiom(ontology, annoPropertyAssertion);
	}


	private static void addShapeProperties(XSLFAutoShape autoShape) {
		
		//Añadimos el nombre de la forma
		 OWLAnnotationProperty property = dataFactory.getOWLAnnotationProperty("title", pm3);
	  	 OWLAnnotation anno = dataFactory.getOWLAnnotation(property, dataFactory.getOWLLiteral(autoShape.getShapeName().toString()));		
	     OWLAnnotationAssertionAxiom annoPropertyAssertion =
	  		dataFactory.getOWLAnnotationAssertionAxiom(IRI.create("http://purl.org/spar/doco/#shape" + nShape), anno);

	     ontManager.addAxiom(ontology, annoPropertyAssertion);
		
	     //Añadimos color de la forma (si lo tiene)
	     if(autoShape.getFillColor() != null) {
	     property = dataFactory.getOWLAnnotationProperty("/color", pm);
	  	 anno = dataFactory.getOWLAnnotation(property, dataFactory.getOWLLiteral(autoShape.getFillColor().toString()));		
	     annoPropertyAssertion =
	  		dataFactory.getOWLAnnotationAssertionAxiom(IRI.create("http://purl.org/spar/doco/#shape" + nShape), anno);

	     ontManager.addAxiom(ontology, annoPropertyAssertion);
	     }
	     //Anadimos la sombra de la forma (si la hay)
	     if(autoShape.getShadow() != null) {
	     property = dataFactory.getOWLAnnotationProperty("/shadow", pm);
	  	 anno = dataFactory.getOWLAnnotation(property, dataFactory.getOWLLiteral(autoShape.getShadow().toString()));		
	     annoPropertyAssertion =
	  		dataFactory.getOWLAnnotationAssertionAxiom(IRI.create("http://purl.org/spar/doco/#shape" + nShape), anno);

	     ontManager.addAxiom(ontology, annoPropertyAssertion);
	     
	     }
	}


	private static void autoshape(XSLFShape shape, OWLNamedIndividual parentSlide) {
		
		XSLFAutoShape autoShape = (XSLFAutoShape) shape;
		if(autoShape.getShapeType() != null)
		if(!autoShape.getShapeType().toString().equals("TEXT_BOX"))
			picture(shape, parentSlide);
				
		OWLNamedIndividual textParent = text(shape, parentSlide);
		
  		if(shape.getPlaceholder() != null) {
		autoShapeType(autoShape,textParent);
		
  		}		
	}
	
	private static void autoShapeType(XSLFAutoShape autoShape, OWLNamedIndividual textParent) {
		
		//aumentamos el nº de autoShapeTypes
		nAutoType++;
		
		String aux = autoShape.getPlaceholder().name();
		OWLClass somethingOnOnt;
		OWLNamedIndividual autoShapeType;
		
		if(aux.equals("CENTERED_TITLE") || aux.equals("TITLE") ) {
			//Añadimos la autoshape a la ontologia como parte de "Title"
	  		somethingOnOnt = dataFactory.getOWLClass(":/Title",pm);
			
		}else if(aux.equals("SUBTITLE")) {
			//Añadimos la textShape a la ontologia como parte de "Subtitle"
			somethingOnOnt = dataFactory.getOWLClass(":/Subtitle",pm);
		}else if(aux.equals("HEADER")) {
			//Añadimos la textShape a la ontologia como parte de "Header"
			somethingOnOnt = dataFactory.getOWLClass(":/Header",pm);
		}else if(aux.equals("FOOTER")) {
			//Añadimos la textShape a la ontologia como parte de "Footer"
			somethingOnOnt = dataFactory.getOWLClass(":/Footnote",pm);
		}else {
			somethingOnOnt = dataFactory.getOWLClass(":/Other",pm);
		}
		
  		   
  		//Se crea la class assertion para indicar que el tipo es una instancia de somethingOnOnt
		autoShapeType = dataFactory.getOWLNamedIndividual(":/#autoShapeType" + nAutoType, pm);
  		OWLClassAssertionAxiom classAssertion = dataFactory.getOWLClassAssertionAxiom(somethingOnOnt, autoShapeType);

  		//Se añade la class assertion a la ontologia
  		ontManager.addAxiom(ontology, classAssertion);

  		//Añadimos el individual como "isContainedBy" de su shape
  		OWLObjectProperty isContainedBy = dataFactory.getOWLObjectProperty("#isContainedBy",pm2);
  		
  		OWLObjectPropertyAssertionAxiom assertion =
  				dataFactory.getOWLObjectPropertyAssertionAxiom(isContainedBy, autoShapeType, textParent);
  		
  		ontManager.addAxiom(ontology, assertion);
	}


	//Method that locates the text contained in a shape and adds it and its characteristics to the ontology
	private static OWLNamedIndividual text(XSLFShape shape, OWLNamedIndividual parent) {
		
		XSLFTextShape textShape = (XSLFTextShape)shape;
		
		//Si hay texto
		if(!textShape.getText().isBlank()) {
		//Si no esta vacio entonces se cuenta el texto	
		nText++;	
        String text = textShape.getText();
       
        //Añadimos la textShape a la ontologia como parte de "TextChunk"
  		OWLClass textOnOnt = dataFactory.getOWLClass(":/TextChunk",pm);
  		   
  		//Se crea la class assertion para indicar que text es una instancia de TextChunk
  		OWLNamedIndividual textChunk = dataFactory.getOWLNamedIndividual(":/#text" + nText, pm);
  		OWLClassAssertionAxiom classAssertion = dataFactory.getOWLClassAssertionAxiom(textOnOnt, textChunk);

  		//Se añade la class assertion a la ontologia
  		ontManager.addAxiom(ontology, classAssertion);

  		//Añadimos el individual como "isContainedBy" de su slide
  		OWLObjectProperty isContainedBy = dataFactory.getOWLObjectProperty("#isContainedBy",pm2);
  		OWLDataProperty data = dataFactory.getOWLDataProperty("description", pm3);
  		
  		OWLDataPropertyAssertionAxiom dataPropertyAssertion =
  								dataFactory.getOWLDataPropertyAssertionAxiom(data, textChunk, text);
  		
  		OWLObjectPropertyAssertionAxiom assertion =
  				dataFactory.getOWLObjectPropertyAssertionAxiom(isContainedBy, textChunk, parent);
  		
  		ontManager.addAxiom(ontology, assertion);
  		ontManager.addAxiom(ontology, dataPropertyAssertion);  		
  		
  		//----------------Añadir parrafos a la ontologia---------
  		paragraphSearch(textShape, textChunk);
  		//-------------------------------------------------------
  		//----------------Añadir listas a la ontologia-----------
  		createList(textShape, textChunk);
  		//-------------------------------------------------------
		
        return textChunk;
        
		}
		
		OWLNamedIndividual textChunk = null;
		return textChunk;
	}

	//This method iterates through the paragraphs present in a textShape and adds them and their components to the ontology
	private static void paragraphSearch(XSLFTextShape textShape, OWLNamedIndividual parent) {
		
		 
		int longParagraph = textShape.getTextParagraphs().size();

		for(int i=0; i < longParagraph; i++) {
			
			//Si el parrafo no esta vacio
			if(!textShape.getTextParagraphs().get(i).getText().isBlank()) {
				
			//aumentamos el nº de parrafos
			nPar++;
			
			//Añadimos el parrafo a la ontologia como parte de "Paragraph" 
	  		OWLClass paragraphOnOnt = dataFactory.getOWLClass(":/Paragraph",pm);
	  		   
	  		//Se crea la class assertion para indicar que paragraph es una instancia de Paragraph
	  		OWLNamedIndividual paragraph = dataFactory.getOWLNamedIndividual(":/#paragraph" + nPar, pm);
	  		OWLClassAssertionAxiom classAssertion = dataFactory.getOWLClassAssertionAxiom(paragraphOnOnt, paragraph);

	  		//Se añade la class assertion a la ontologia
	  		ontManager.addAxiom(ontology, classAssertion);
			
	  		//Añadimos el individual como "isContainedBy" de su texto
	  		OWLObjectProperty isContainedBy = dataFactory.getOWLObjectProperty("#isContainedBy",pm2);
	  		
	  		OWLObjectPropertyAssertionAxiom assertion =
	  				dataFactory.getOWLObjectPropertyAssertionAxiom(isContainedBy, paragraph, parent);
	  		
	  		ontManager.addAxiom(ontology, assertion);
	  		
	  		//Añadimos propiedades del parrafo
	  		addParagraphProperties(textShape.getTextParagraphs().get(i));
	  		
	  		//Busqueda de hipervinculos
	  		hyperlinkSearch(textShape.getTextParagraphs().get(i),paragraph);
	  		
	  		//Busqueda tipos de letra
	  		textTypesSearch(textShape.getTextParagraphs().get(i));

		}	
		
	}

	}

	private static void addParagraphProperties(XSLFTextParagraph parrafo) {
		
		 //Añadimos el texto al parrafo
		 OWLAnnotationProperty property = dataFactory.getOWLAnnotationProperty("description", pm3);
	  	 OWLAnnotation anno = dataFactory.getOWLAnnotation(property, dataFactory.getOWLLiteral(parrafo.getText().toString()));		
	     OWLAnnotationAssertionAxiom annoPropertyAssertion =
	  		dataFactory.getOWLAnnotationAssertionAxiom(IRI.create("http://purl.org/spar/doco/#paragraph" + nPar), anno);

	     ontManager.addAxiom(ontology, annoPropertyAssertion);
	     
	     //Añadimos alineamiento del parrafo
	     property = dataFactory.getOWLAnnotationProperty("/alignment", pm);
	  	 anno = dataFactory.getOWLAnnotation(property, dataFactory.getOWLLiteral(parrafo.getTextAlign().toString()));		
	     annoPropertyAssertion =
	  		dataFactory.getOWLAnnotationAssertionAxiom(IRI.create("http://purl.org/spar/doco/#paragraph" + nPar), anno);

	     ontManager.addAxiom(ontology, annoPropertyAssertion);
	     
	     //Añadimos interlineado del parrafo (si existe)
	     if(parrafo.getLineSpacing() != null) { 
	     property = dataFactory.getOWLAnnotationProperty("/lineSpacing", pm);
	  	 anno = dataFactory.getOWLAnnotation(property, dataFactory.getOWLLiteral(parrafo.getLineSpacing()/100));		
	     annoPropertyAssertion =
	  		dataFactory.getOWLAnnotationAssertionAxiom(IRI.create("http://purl.org/spar/doco/#paragraph" + nPar), anno);

	     ontManager.addAxiom(ontology, annoPropertyAssertion);
	     }
		
	}

	//This method finds the different formats and types contained in a paragraph and adds them to the ontology
	private static void textTypesSearch(XSLFTextParagraph xslfTextParagraph) {
		 	
		int nRuns =  xslfTextParagraph.getTextRuns().size();
		for(int i = 0; i < nRuns; i++ ) {
			
			//Añadimos estas propiedades al parrafo en el que se encuentran
			addTextTypeProperties(xslfTextParagraph.getTextRuns().get(i));
			
		}	
			

		
	}


	private static void addTextTypeProperties(XSLFTextRun xslfTextRun) {
		
		//Añadimos el tamaño de la fuente
		 OWLAnnotationProperty property = dataFactory.getOWLAnnotationProperty("/fontSize", pm);
	  	 OWLAnnotation anno = dataFactory.getOWLAnnotation(property, dataFactory.getOWLLiteral(xslfTextRun.getFontSize().toString()));		
	     OWLAnnotationAssertionAxiom annoPropertyAssertion =
	  		dataFactory.getOWLAnnotationAssertionAxiom(IRI.create("http://purl.org/spar/doco/#paragraph" + nPar), anno);

	     ontManager.addAxiom(ontology, annoPropertyAssertion);
	     
	    //Añadimos el tipo de la fuente
		 property = dataFactory.getOWLAnnotationProperty("/fontFamily", pm);
	  	 anno = dataFactory.getOWLAnnotation(property, dataFactory.getOWLLiteral(xslfTextRun.getFontFamily()));		
	     annoPropertyAssertion =
	  		dataFactory.getOWLAnnotationAssertionAxiom(IRI.create("http://purl.org/spar/doco/#paragraph" + nPar), anno);

	     ontManager.addAxiom(ontology, annoPropertyAssertion);
	     
	   //Añadimos el color de la fuente
		 property = dataFactory.getOWLAnnotationProperty("/color", pm);
	  	 anno = dataFactory.getOWLAnnotation(property, dataFactory.getOWLLiteral(xslfTextRun.getFontColor().toString()));		
	     annoPropertyAssertion =
	  		dataFactory.getOWLAnnotationAssertionAxiom(IRI.create("http://purl.org/spar/doco/#paragraph" + nPar), anno);

	     ontManager.addAxiom(ontology, annoPropertyAssertion);
	     
	   //Añadimos modificadores de la fuente
		 property = dataFactory.getOWLAnnotationProperty("/textModifier", pm);
		 
		 if(xslfTextRun.isItalic()) {
		  	 anno = dataFactory.getOWLAnnotation(property, dataFactory.getOWLLiteral("Italic: " + xslfTextRun.getRawText()));		
		  	annoPropertyAssertion =
			  		dataFactory.getOWLAnnotationAssertionAxiom(IRI.create("http://purl.org/spar/doco/#paragraph" + nPar), anno);
		  	ontManager.addAxiom(ontology, annoPropertyAssertion);
		  	
		 }
		 
		 if(xslfTextRun.isBold()) {
		  	 anno = dataFactory.getOWLAnnotation(property, dataFactory.getOWLLiteral("Bold: " + xslfTextRun.getRawText()));		
		  	annoPropertyAssertion =
			  		dataFactory.getOWLAnnotationAssertionAxiom(IRI.create("http://purl.org/spar/doco/#paragraph" + nPar), anno);
		  	ontManager.addAxiom(ontology, annoPropertyAssertion);
		  	
		 }
		 
		 if(xslfTextRun.isStrikethrough()) {
		  	 anno = dataFactory.getOWLAnnotation(property, dataFactory.getOWLLiteral("Strikethrough: " + xslfTextRun.getRawText()));		
		  	annoPropertyAssertion =
			  		dataFactory.getOWLAnnotationAssertionAxiom(IRI.create("http://purl.org/spar/doco/#paragraph" + nPar), anno);
		  	ontManager.addAxiom(ontology, annoPropertyAssertion);
		  	
		 }
		 
		 if(xslfTextRun.isUnderlined()) {
		  	 anno = dataFactory.getOWLAnnotation(property, dataFactory.getOWLLiteral("Underlined: " + xslfTextRun.getRawText()));		
		  	annoPropertyAssertion =
			  		dataFactory.getOWLAnnotationAssertionAxiom(IRI.create("http://purl.org/spar/doco/#paragraph" + nPar), anno);
		  	ontManager.addAxiom(ontology, annoPropertyAssertion);
		 }		
	}


	//This method searches a paragraph´s text runs to find hyperlinks and adds them to the ontology
	private static void hyperlinkSearch(XSLFTextParagraph xslfTextParagraph, OWLNamedIndividual parent) {
		

		int longTextRun =  xslfTextParagraph.getTextRuns().size() ;
		
		for(int j=0; j < longTextRun; j++) {
			
			if(xslfTextParagraph.getTextRuns().get(j).getHyperlink() != null) {
				
				//Aumentamos el numero de hypervinculos del ppt
				nLink++;

				
				 //Añadimos el hyperlink a la ontologia como parte de "Hyperlink" (no existe)
		  		OWLClass linkOnOnt = dataFactory.getOWLClass(":/Hyperlink",pm);
		  		   
		  		//Se crea la class assertion para indicar que link es una instancia de Hyperlink
		  		OWLNamedIndividual hyperlink = dataFactory.getOWLNamedIndividual(":/#hyperlink" + nLink, pm);
		  		
		  		OWLClassAssertionAxiom classAssertion = dataFactory.getOWLClassAssertionAxiom(linkOnOnt, hyperlink);
		  			
		  		//Se añade la class assertion a la ontologia
		  		ontManager.addAxiom(ontology, classAssertion);
				
		  		//Añadimos el individual como "isContainedBy" de su parrafo
		  		OWLObjectProperty isContainedBy = dataFactory.getOWLObjectProperty("#isContainedBy",pm2);
		  		
		  		OWLObjectPropertyAssertionAxiom assertion =
		  				dataFactory.getOWLObjectPropertyAssertionAxiom(isContainedBy, hyperlink, parent);
		  		
		  		ontManager.addAxiom(ontology, assertion);
		  		
		  		//Añadimos las propiedades del vinculo
		  		addHyperlinkProperties(xslfTextParagraph.getTextRuns().get(j));
		  			
			}
		}
	}

	//This method adds the properties of a hyperlink to the ontology
	private static void addHyperlinkProperties(XSLFTextRun textRun) {
		
		 //Añadimos el texto del link
		 OWLAnnotationProperty property = dataFactory.getOWLAnnotationProperty("description", pm3);
	  	 OWLAnnotation anno = dataFactory.getOWLAnnotation(property, dataFactory.getOWLLiteral(textRun.getRawText()));		
	     OWLAnnotationAssertionAxiom annoPropertyAssertion =
	  		dataFactory.getOWLAnnotationAssertionAxiom(IRI.create("http://purl.org/spar/doco/#hyperlink" + nLink), anno);

	     ontManager.addAxiom(ontology, annoPropertyAssertion);
	     
	     //Añadimos la URL del link
	     property = dataFactory.getOWLAnnotationProperty("/linkAddress", pm);
	  	 anno = dataFactory.getOWLAnnotation(property, dataFactory.getOWLLiteral(textRun.getHyperlink().getAddress()));		
	     annoPropertyAssertion =
	  		dataFactory.getOWLAnnotationAssertionAxiom(IRI.create("http://purl.org/spar/doco/#hyperlink" + nLink), anno);

	     ontManager.addAxiom(ontology, annoPropertyAssertion);
		
	}


	private static void createList(XSLFTextShape textShape, OWLNamedIndividual parent) {
		
		

			//Si hay una list en el texto
			if(isThereList(textShape)) {
				
			//Creacion de una lista nueva en la ontologia	
				//Aumentamos el numero de listas del ppt
				nList++;
				
				 //Añadimos la lista a la ontologia como parte de "List" 
		  		OWLClass listOnOnt = dataFactory.getOWLClass(":/List",pm5);
		  		   
		  		//Se crea la class assertion para indicar que list es una instancia de List
		  		OWLNamedIndividual list = dataFactory.getOWLNamedIndividual(":/#list" + nList, pm5);
		  		OWLClassAssertionAxiom classAssertion = dataFactory.getOWLClassAssertionAxiom(listOnOnt, list);

		  		//Se añade la class assertion a la ontologia
		  		ontManager.addAxiom(ontology, classAssertion);

		  		//Añadimos el individual lista como "isContainedBy" de su shape
		  		OWLObjectProperty isContainedBy = dataFactory.getOWLObjectProperty("#isContainedBy",pm2);
		  		
			     
		  		OWLObjectPropertyAssertionAxiom assertion =
		  				dataFactory.getOWLObjectPropertyAssertionAxiom(isContainedBy, list, parent);
		  		
		  		ontManager.addAxiom(ontology, assertion);
				
				
				//Añadir elementos a la lista por parrafos
				addListElements(textShape, list);
			}		
	}

	private static void addListElements(XSLFTextShape textShape, OWLNamedIndividual list) {
		
		int auxNElems = 0;
		int longParagraph = textShape.getTextParagraphs().size();
		for(int i=0; i < longParagraph; i++) {
			
			//si hay un elemento de una lista en un parrafo se añade a la lista
			if(textShape.getTextParagraphs().get(i).isBullet()) {	
				
				//Creacion de un listElement nuevo en la ontologia	
				//Aumentamos el numero de listElems del ppt
				nListElems++;
				
				 //Añadimos la lista a la ontologia como parte de "ListItem"
		  		OWLClass listElemOnOnt = dataFactory.getOWLClass(":/ListItem",pm5);
		  		   
		  		//Se crea la class assertion para indicar que list es una instancia de List
		  		OWLNamedIndividual listElement =
		  				dataFactory.getOWLNamedIndividual(":/#listItem" + nListElems, pm5);
		  		OWLClassAssertionAxiom classAssertion =
		  				dataFactory.getOWLClassAssertionAxiom(listElemOnOnt, listElement);

		  		//Se añade la class assertion a la ontologia
		  		ontManager.addAxiom(ontology, classAssertion);

				//Añadimos el individual listElement como "isContainedBy" de su lista
		  		OWLObjectProperty isContainedBy =
		  				dataFactory.getOWLObjectProperty("#isContainedBy",pm2);
		  		
		  		OWLObjectPropertyAssertionAxiom assertion =
		  		dataFactory.getOWLObjectPropertyAssertionAxiom(isContainedBy, listElement, list);
		  		
		  		ontManager.addAxiom(ontology, assertion);
		  		
		  		//Añadimos el texto al listItem
				 OWLAnnotationProperty property = 
						 dataFactory.getOWLAnnotationProperty("description", pm3);
			  	 OWLAnnotation anno = 
			  			 dataFactory.getOWLAnnotation(property, 
			  					 dataFactory.getOWLLiteral(textShape.getTextParagraphs().get(i).getText()));		
			     OWLAnnotationAssertionAxiom annoPropertyAssertion =
			  		dataFactory.getOWLAnnotationAssertionAxiom
			  		(IRI.create("http://purl.org/spar/doco/#listItem" + nListElems), anno);

			     ontManager.addAxiom(ontology, annoPropertyAssertion);

		  	  //Aumentamos el numero de elems para poner en la lista
		  		auxNElems++;
			}
		}				
		//Añadimos el nº de elementos como annotation de su lista
		 OWLAnnotationProperty property =
				 dataFactory.getOWLAnnotationProperty("/listSize", pm);
	  	 OWLAnnotation anno =
	  			 dataFactory.getOWLAnnotation(property, dataFactory.getOWLLiteral(auxNElems));		
	     OWLAnnotationAssertionAxiom annoPropertyAssertion =
	  		dataFactory.getOWLAnnotationAssertionAxiom
	  		(IRI.create("http://purl.org/spar/doco/#list" + nList), anno);

	     ontManager.addAxiom(ontology, annoPropertyAssertion);
	     
	}



	//Returns true if there is at least one element of the type "bullet"
	private static boolean isThereList(XSLFTextShape textShape) {
		
		int longParagraph = textShape.getTextParagraphs().size();
		boolean found = false;
		for(int i=0; i < longParagraph && !found; i++) {
			
			//si hay un elemento de una lista en un parrafo se sale y devuelve true
			if(textShape.getTextParagraphs().get(i).isBullet() && (!textShape.getTextParagraphs().get(i).getText().isBlank()) ) {	
				found = true;
			}
		}
		return found;
	}



	private static int find(String[] typeList, String obj) {
		
		int index = 0;
		for ( int i =0; i < typeList.length; i++) {
			
			if ( typeList[i].equals(obj) ) {
				index = i;
				break;
			}
		}
		
		return index;
	}
	
	
	private static boolean contains(String[] typeList, String obj) {
		
		boolean found = false;
		int n =0;
		while ( (n < typeList.length) && !found ) {
			
			if ( typeList[n].equals(obj) ) {
				found = true;
			}
			else {
				n++;
			}
		}
		
		return found;
	}

	private static String nombreTipo(String string) {
		int numDots = 0;
		for (int j = 0; j < string.length(); j++){
			
			if (string.charAt(j) == '.') {
				numDots++;
			}
		}
		String nombre = "";
		int countAux = 0;
		int flag = 0;
		
		for (int i = 0; i < string.length(); i++) {
			
			if (string.charAt(i) == '.') {
				countAux++;
			}
			if(flag == 1) {
				nombre += string.charAt(i);
			}
			if(countAux == numDots) {
				flag = 1;
			}
			
		}
		return nombre;
	}	

}
