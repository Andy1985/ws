JAVA_HOME = /usr/java/jdk1.5.0_11
LIB_DIR = ./lib

CLASSPATH = $(LIB_DIR)/slf4j-api-1.4.3.jar:$(LIB_DIR)/mina-core-1.1.7.jar:$(LIB_DIR)/poi-scratchpad-3.6-20091214.jar:$(LIB_DIR)/slf4j-log4j12-1.4.3.jar:$(LIB_DIR)/poi-3.9-20121203.jar:$(LIB_DIR)/poi-scratchpad-3.9-20121203.jar:$(LIB_DIR)/poi-ooxml-3.9-20121203.jar:$(LIB_DIR)/poi-ooxml-schemas-3.9-20121203.jar:$(LIB_DIR)/poi-excelant-3.9-20121203.jar:$(LIB_DIR)/xmlbeans-2.3.0.jar:$(LIB_DIR)/stax-api-1.0.1.jar:$(LIB_DIR)/dom4j-1.6.1.jar:$(LIB_DIR)/pdfbox-app-1.8.2.jar:$(LIB_DIR)/java-unrar-0.3.jar:$(LIB_DIR)/ant.jar


TARGET = ws.jar

all: $(TARGET)

$(TARGET): WordServer.java WordServerHandler.java Convertor.java
	$(JAVA_HOME)/bin/javac WordServer.java WordServerHandler.java Convertor.java -classpath $(CLASSPATH)
	$(JAVA_HOME)/bin/jar cvfm ws.jar MANIFEST.MF WordServer.class 
	$(JAVA_HOME)/bin/jar cvf wordserver.jar log4j.properties WordServerHandler.class Convertor.class WordServerHandler\$$Worker.class
	cp ws.jar $(LIB_DIR) -f
	cp wordserver.jar $(LIB_DIR) -f

.PHONY: clean
clean:
	rm -f *.class
	rm -f *.jar
