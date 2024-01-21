package com.sanaa.powerpoint.model;


import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;
import org.apache.poi.xslf.usermodel.XSLFTextBox;

import java.io.FileOutputStream;
import java.io.IOException;

;
public class ITMastersStudentPresentation {
	
	public static void main(String[] args) {
        try {
            // Create a new PowerPoint presentation
            XMLSlideShow ppt = new XMLSlideShow();

            // Slide 1: Introduction
            XSLFSlide slide1 = ppt.createSlide();
            setSlideTitleAndContent(slide1, "IT Master's Student", "Hello, I am an enthusiastic IT master's student passionate about technology and innovation.");

            // Slide 2: Academic Background
            XSLFSlide slide2 = ppt.createSlide();
            setSlideTitleAndContent(slide2, "Academic Background", "Currently pursuing a Master's in Information Technology at XYZ University, with expected graduation in 2023.");

            // Slide 3: Skills and Projects
            XSLFSlide slide3 = ppt.createSlide();
            setSlideTitleAndContent(slide3, "Skills and Projects", "Proficient in programming languages such as Python and Java. Completed projects include a web-based e-commerce application and a machine learning model for data analysis.");

            // Save the presentation to a file
            try (FileOutputStream out = new FileOutputStream("IT_Masters_Student_Presentation.pptx")) {
                ppt.write(out);
            }

            System.out.println("Presentation created successfully.");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    // Helper method to set title and content on a slide
    private static void setSlideTitleAndContent(XSLFSlide slide, String title, String content) {
        XSLFTextBox titleTextBox = slide.createTextBox();
        titleTextBox.setText(title);
        titleTextBox.setAnchor(new java.awt.Rectangle(50, 50, 600, 50));

        XSLFTextBox contentTextBox = slide.createTextBox();
        contentTextBox.setText(content);
        contentTextBox.setAnchor(new java.awt.Rectangle(50, 100, 600, 300));
    }

}
