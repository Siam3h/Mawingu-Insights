import streamlit as st
import plotly.express as px
import pandas as pd 
import warnings
from wordcloud import WordCloud
import matplotlib.pyplot as plt
import time

st.set_page_config(layout="wide")
warnings.filterwarnings('ignore')

st.markdown("<h2 style='text-align: center; padding-top: 1rem;'>Study of Studies</h2>", unsafe_allow_html=True)

fl = st.file_uploader(":file_folder: Upload a file",type=(["csv"]))

if fl is not None:
    
    # Read the CSV file
    df = pd.read_csv(fl, encoding="ISO-8859-1")

    page = st.sidebar.selectbox("Select a page", ["Home","Word Cloud"])

    # Flash success message for 5 seconds
    success_message = st.empty()
    success_message.success("File successfully uploaded!")
    time.sleep(3)
    success_message.empty()

    #Sidebar for Document Type
    st.sidebar.header("Choose your filter: ")

    doctype = st.sidebar.multiselect("Document Type", df["DOCUMENT TYPE (book, Journal article, Report, Statutes, Newspaper)"].unique())
    if not doctype:
        df2 = df.copy()
    else:
        df2 = df[df["DOCUMENT TYPE (book, Journal article, Report, Statutes, Newspaper)"].isin(doctype)]


    #sidebar for Applicable Nomenclature
    applicableNom = st.sidebar.multiselect("Applicable Nomenclature",df2["APPLICABLE NOMENCLATURE(KEY NAME USED)"].unique())
    if not applicableNom:
        df3 = df2.copy()
    else:
        df3 = df2[df2["APPLICABLE NOMENCLATURE(KEY NAME USED)"].isin(applicableNom)]

    #sidebar for Section of Article Analysed
    sectionAnalysed = st.sidebar.multiselect("Section of Article Analysed", df3["SECTION OF THE ARTICLE ANALYSED"].unique())
    if not sectionAnalysed:
        df4 = df3.copy()
    else:
        df4=df3[df3["SECTION OF THE ARTICLE ANALYSED"].isin(sectionAnalysed)]

    #sidebar for Nationality of Author
    nationality = st.sidebar.multiselect("Nationality of Author", df4["NATIONALITY/ORIGIN/EXTRACTION OF AUTHOR"].unique())
    if not nationality:
        df5 = df4.copy()
    else:
        df5 = df4[df4["NATIONALITY/ORIGIN/EXTRACTION OF AUTHOR"].isin(nationality)]

    #sidebar for Year of Publication
    yearofPub = st.sidebar.multiselect("Year of Publication", df5["YEAR OF PUBLICATION"].unique())
    if not yearofPub:
        df6 = df5.copy()
    else:
        df6 = df5[df5["YEAR OF PUBLICATION"].isin(yearofPub)]

    #sidebar for Positionality of Articles
    positionality = st.sidebar.multiselect("Positionality of Article", df6["POSITIONALITY OF ARTICLE"].unique())
    if not positionality:
        df7 = df6.copy()
    else:
        df7 = df6[df6["POSITIONALITY OF ARTICLE"].isin(positionality)]

    #sidebar for Categorization of Publication
    categorization = st.sidebar.multiselect("Categorization of Publication", df7["CATEGORISATION OF PUBLICATION( descriptive, diagnostic(identify problems), Prescriptive, theoretical)"].unique())
    if not categorization:
        df8 = df7.copy()
    else:
        df8 = df7[df7["CATEGORISATION OF PUBLICATION( descriptive, diagnostic(identify problems), Prescriptive, theoretical)"].isin(categorization)]

    #sidebar for Continent of Study
    #continent = st.sidebar.multiselect("Continent of Study", df8["CONTINENT OF STUDY"].unique())
    #if not continent:
     #   df9 = df8.copy()
    #else:
     #   df9 = df8[df8["CONTINENT OF STUDY"].isin(continent)]

    #sidebar for Use of Statistics
    stats = st.sidebar.multiselect("Use of Statistics", df8["USE OF STATISTICS"].unique())
    if not stats:
        df9 = df8.copy()
    else:
        df9 = df8[df8["USE OF STATISTICS"].isin(stats)]

    #sidebar for CITE 'UN-habitat 2003 STUDY ON SLUMS'
    un_citation = st.sidebar.multiselect("UN-Habitat Citation", df9["CITE 'UN-habitat 2003 STUDY ON SLUMS'"].unique())
    if not un_citation:
        df10 = df9.copy()
    else:
        df10 = df9[df9["CITE 'UN-habitat 2003 STUDY ON SLUMS'"].isin(un_citation)]
     
    #Filter the data based on the DF.
    if not doctype and not stats and not un_citation and not applicableNom and not sectionAnalysed and not nationality and not yearofPub and not positionality and not categorization and not stats and not un_citation:
        filtered_df = df

    elif not doctype and not applicableNom and not sectionAnalysed and not nationality and not yearofPub and not positionality  and not stats and not un_citation:
        filtered_df = df[df["CATEGORISATION OF PUBLICATION( descriptive, diagnostic(identify problems), Prescriptive, theoretical)"].isin(categorization)]
    elif not doctype and not applicableNom and not sectionAnalysed and not nationality and not yearofPub and not categorization  and not stats and not un_citation :
        filtered_df = df[df["POSITIONALITY OF ARTICLE"].isin(positionality)]
    elif not doctype and not applicableNom and not sectionAnalysed and not nationality and not positionality and not categorization  and not stats and not un_citation :
        filtered_df = df[df["YEAR OF PUBLICATION"].isin(yearofPub)]
    elif not doctype and not applicableNom and not sectionAnalysed and not yearofPub and not positionality and not categorization  and not stats and not un_citation :
        filtered_df = df[df["NATIONALITY/ORIGIN/EXTRACTION OF AUTHOR"].isin(nationality)]
    elif not doctype and not applicableNom and not nationality and not yearofPub and not positionality and not categorization  and not stats and not un_citation :
        filtered_df = df[df["SECTION OF THE ARTICLE ANALYSED"].isin(sectionAnalysed)]
    elif not doctype and not sectionAnalysed and not nationality and not yearofPub and not positionality and not categorization  and not stats and not un_citation :
        filtered_df = df[df["APPLICABLE NOMENCLATURE(KEY NAME USED)"].isin(applicableNom)]
    elif not applicableNom and not sectionAnalysed and not nationality and not yearofPub and not positionality and not categorization  and not stats and not un_citation :
        filtered_df = df[df["DOCUMENT TYPE (book, Journal article, Report, Statutes, Newspaper)"].isin(doctype)]
    elif not applicableNom and not sectionAnalysed and not nationality and not yearofPub and not positionality and not categorization and not stats:
        filtered_df = df[df["CITE 'UN-habitat 2003 STUDY ON SLUMS'"].isin(un_citation)]
    elif not applicableNom and not sectionAnalysed and not nationality and not yearofPub and not positionality and not categorization and not un_citation:
        filtered_df = df[df["USE OF STATISTICS"].isin(stats)]

    elif doctype and applicableNom:
        filtered_df = df3[df["DOCUMENT TYPE (book, Journal article, Report, Statutes, Newspaper)"].isin(doctype) & df3["APPLICABLE NOMENCLATURE(KEY NAME USED)"].isin(applicableNom)]
    elif doctype and sectionAnalysed:
        filtered_df = df3[df["DOCUMENT TYPE (book, Journal article, Report, Statutes, Newspaper)"].isin(doctype) & df3["SECTION OF THE ARTICLE ANALYSED"].isin(sectionAnalysed)]
    elif doctype and nationality:
        filtered_df = df3[df["DOCUMENT TYPE (book, Journal article, Report, Statutes, Newspaper)"].isin(doctype) & df3["NATIONALITY/ORIGIN/EXTRACTION OF AUTHOR"].isin(nationality)]
    elif doctype and yearofPub:
        filtered_df = df3[df["DOCUMENT TYPE (book, Journal article, Report, Statutes, Newspaper)"].isin(doctype) & df3["YEAR OF PUBLICATION"].isin(yearofPub)]
    elif doctype and positionality:
        filtered_df = df3[df["DOCUMENT TYPE (book, Journal article, Report, Statutes, Newspaper)"].isin(doctype) & df3["POSITIONALITY OF ARTICLE"].isin(positionality)]
    elif doctype and categorization:
        filtered_df = df3[df["DOCUMENT TYPE (book, Journal article, Report, Statutes, Newspaper)"].isin(doctype) & df3["CATEGORISATION OF PUBLICATION( descriptive, diagnostic(identify problems), Prescriptive, theoretical)"].isin(categorization)]

    elif applicableNom and nationality:
        filtered_df = df3[df["APPLICABLE NOMENCLATURE(KEY NAME USED)"].isin(applicableNom) & df3["NATIONALITY/ORIGIN/EXTRACTION OF AUTHOR"].isin(nationality)]
    elif applicableNom and sectionAnalysed:
        filtered_df = df3[df["APPLICABLE NOMENCLATURE(KEY NAME USED)"].isin(applicableNom) & df3["SECTION OF THE ARTICLE ANALYSED"].isin(sectionAnalysed)]
    elif applicableNom and yearofPub:
        filtered_df = df3[df["APPLICABLE NOMENCLATURE(KEY NAME USED)"].isin(applicableNom) & df3["YEAR OF PUBLICATION"].isin(yearofPub)]
    elif applicableNom and positionality:
        filtered_df = df3[df["APPLICABLE NOMENCLATURE(KEY NAME USED)"].isin(applicableNom) & df3["POSITIONALITY OF ARTICLE"].isin(positionality)]
    elif applicableNom and categorization:
        filtered_df = df3[df["APPLICABLE NOMENCLATURE(KEY NAME USED)"].isin(applicableNom) & df3["CATEGORISATION OF PUBLICATION( descriptive, diagnostic(identify problems), Prescriptive, theoretical)"].isin(categorization)]
    

    elif nationality and sectionAnalysed:
        filtered_df = df3[df["NATIONALITY/ORIGIN/EXTRACTION OF AUTHOR"].isin(nationality) & df3["SECTION OF THE ARTICLE ANALYSED"].isin(sectionAnalysed)]
    elif nationality and yearofPub:
        filtered_df = df3[df["NATIONALITY/ORIGIN/EXTRACTION OF AUTHOR"].isin(nationality) & df3["YEAR OF PUBLICATION"].isin(yearofPub)]
    elif nationality and positionality:
        filtered_df = df3[df["NATIONALITY/ORIGIN/EXTRACTION OF AUTHOR"].isin(nationality) & df3["POSITIONALITY OF ARTICLE"].isin(positionality)]
    elif nationality and categorization:
        filtered_df = df3[df["NATIONALITY/ORIGIN/EXTRACTION OF AUTHOR"].isin(nationality) & df3["CATEGORISATION OF PUBLICATION( descriptive, diagnostic(identify problems), Prescriptive, theoretical)"].isin(categorization)]


    elif sectionAnalysed and yearofPub:
        filtered_df = df3[df["SECTION OF THE ARTICLE ANALYSED"].isin(sectionAnalysed) & df3["YEAR OF PUBLICATION"].isin(yearofPub)]
    elif sectionAnalysed and positionality:
        filtered_df = df3[df["SECTION OF THE ARTICLE ANALYSED"].isin(sectionAnalysed) & df3["POSITIONALITY OF ARTICLE"].isin(positionality)]
    elif sectionAnalysed and categorization:
        filtered_df = df3[df["SECTION OF THE ARTICLE ANALYSED"].isin(sectionAnalysed) & df3["CATEGORISATION OF PUBLICATION( descriptive, diagnostic(identify problems), Prescriptive, theoretical)"].isin(categorization)]


    elif yearofPub and positionality:
        filtered_df = df3[df["YEAR OF PUBLICATION"].isin(yearofPub) & df3["POSITIONALITY OF ARTICLE"].isin(positionality)]
    elif yearofPub and categorization:
        filtered_df = df3[df["YEAR OF PUBLICATION"].isin(yearofPub) & df3["CATEGORISATION OF PUBLICATION( descriptive, diagnostic(identify problems), Prescriptive, theoretical)"].isin(categorization)]


    elif positionality and categorization:
        filtered_df = df3[df["POSITIONALITY OF ARTICLE"].isin(positionality) & df3["CATEGORISATION OF PUBLICATION( descriptive, diagnostic(identify problems), Prescriptive, theoretical)"].isin(categorization)]
    
    elif positionality and categorization and doctype:
        filtered_df = df3[df["POSITIONALITY OF ARTICLE"].isin(positionality) & df3["CATEGORISATION OF PUBLICATION( descriptive, diagnostic(identify problems), Prescriptive, theoretical)"].isin(categorization) & df3["DOCUMENT TYPE (book, Journal article, Report, Statutes, Newspaper)"].isin(doctype)]


    else:
        filtered_df = df3[df3["DOCUMENT TYPE (book, Journal article, Report, Statutes, Newspaper)"].isin(doctype) &
                        df3["NATIONALITY/ORIGIN/EXTRACTION OF AUTHOR"].isin(nationality) & 
                        df3["APPLICABLE NOMENCLATURE(KEY NAME USED)"].isin(applicableNom) & 
                        df3["POSITIONALITY OF ARTICLE"].isin(positionality)  & 
                        df3["SECTION OF THE ARTICLE ANALYSED"].isin(sectionAnalysed) & 
                        df3["YEAR OF PUBLICATION"].isin(yearofPub) & 
                            df3["CATEGORISATION OF PUBLICATION( descriptive, diagnostic(identify problems), Prescriptive, theoretical)"].isin(categorization)]

    if page == "Home":
        
        st.markdown("<p style='text-align: center; padding: 1rem; color: Green;'>This is the home page.<br/> Use the sidebar to navigate between pages.</p>", unsafe_allow_html=True)

        ## Check if the filtered DataFrame is not empty before proceeding with visualizations
        if not filtered_df.empty:
            # Display the filtered data
            st.markdown("<h3 style='text-align: center;'>Filtered Data</h3>", unsafe_allow_html=True)
            st.write(filtered_df)

            st.markdown("<h2 style='text-align: center; padding-top: 1rem; padding-bottom: 2rem;'>Column Graphs</h2>", unsafe_allow_html=True)

            col1, col2 = st.columns(2)

            # Visualization 1 - DOCUMENT TYPE (Bar Graph)
            fig1 = px.bar(filtered_df['DOCUMENT TYPE (book, Journal article, Report, Statutes, Newspaper)'].value_counts(), x=filtered_df['DOCUMENT TYPE (book, Journal article, Report, Statutes, Newspaper)'].value_counts().index, y=filtered_df['DOCUMENT TYPE (book, Journal article, Report, Statutes, Newspaper)'].value_counts().values, labels={'x': 'DOCUMENT TYPE (book, Journal article, Report, Statutes, Newspaper)', 'y': 'Count'}, title='Document Type Distribution')
            fig1.update_layout(height=500, width=600)
            fig1.update_layout(margin=dict(l=20, r=20))
            fig1.update_layout(title=dict(x=0.5, y=1, xanchor='center', yanchor='top'))
            col1.plotly_chart(fig1)

            # Visualization 2 - APPLICABLE NOMENCLATURE (Bar Graph)
            fig2 = px.bar(filtered_df['APPLICABLE NOMENCLATURE(KEY NAME USED)'].value_counts(), x=filtered_df['APPLICABLE NOMENCLATURE(KEY NAME USED)'].value_counts().index, y=filtered_df['APPLICABLE NOMENCLATURE(KEY NAME USED)'].value_counts().values, labels={'x': 'APPLICABLE NOMENCLATURE(KEY NAME USED)', 'y': 'Count'}, title='Nomenclature Distribution')
            fig2.update_layout(height=500, width=550)
            fig2.update_layout(margin=dict(l=155, r=10))
            fig2.update_layout(title=dict(x=0.5, y=1, xanchor='center', yanchor='top'))
            col2.plotly_chart(fig2)

            # Visualization 3 - YEAR OF PUBLICATION (Line Chart)
            fig3 = px.line(df['YEAR OF PUBLICATION'].value_counts().sort_index(), labels={'x': 'Year of Publication', 'y': 'Count'},
                        title='Year of Publication Distribution', markers=True)
            
            fig3.update_layout(height=600, width=1000 )
            fig3.update_layout(margin=dict(l=20, r=30))
            fig3.update_layout(title=dict(x=0.5, y=1, xanchor='center', yanchor='top'))

            # Remove white backgrounds
            fig3.update_layout({
                'plot_bgcolor': 'rgba(0, 0, 0, 0)',
                'paper_bgcolor': 'rgba(0, 0, 0, 0)',
            })

            # Visualization 6 - SECTION OF THE ARTICLE ANALYSED (Bar Graph)
            fig6 = px.bar(filtered_df['SECTION OF THE ARTICLE ANALYSED'].value_counts(), x=filtered_df['SECTION OF THE ARTICLE ANALYSED'].value_counts().index, y=filtered_df['SECTION OF THE ARTICLE ANALYSED'].value_counts().values, labels={'x': 'Section', 'y': 'Count'}, title='Section Distribution')
            fig6.update_layout(height=500, width=550)
            fig6.update_layout(margin=dict(l=155, r=10))
            fig6.update_layout(title=dict(x=0.5, y=1, xanchor='center', yanchor='top'))
            col2.plotly_chart(fig6)

            # Visualization 7 - POSITIONALITY OF ARTICLE (Bar Graph)
            fig7 = px.bar(filtered_df['POSITIONALITY OF ARTICLE'].value_counts(), x=filtered_df['POSITIONALITY OF ARTICLE'].value_counts().index, y=filtered_df['POSITIONALITY OF ARTICLE'].value_counts().values, labels={'x': 'Positionality', 'y': 'Count'}, title='Positionality Distribution')
            fig7.update_layout(height=500, width=600)
            fig7.update_layout(margin=dict(l=20, r=20))
            fig7.update_layout(title=dict(x=0.5, y=1, xanchor='center', yanchor='top'))
            col1.plotly_chart(fig7)

            # Visualization 4 - CONTINENT OF STUDY (Pie Chart)
            #fig4 = px.pie(filtered_df, names='CONTINENT OF STUDY', title='Continent of Study Distribution')
            #fig4.update_layout(height=500, width=500)
            #fig4.update_layout(margin=dict(l=20, r=20))
            #fig4.update_layout(title=dict(x=0.5, y=1, xanchor='center', yanchor='top'))
            #col2.plotly_chart(fig4)

            # Visualization 5 - NATIONALITY/ORIGIN/EXTRACTION OF AUTHOR (Pie Chart)
            fig5 = px.pie(filtered_df, names='NATIONALITY/ORIGIN/EXTRACTION OF AUTHOR', title='Nationality Distribution')
            fig5.update_layout(height=500, width=500)
            fig5.update_layout(margin=dict(l=20, r=20))
            fig5.update_layout(title=dict(x=0.5, y=1, xanchor='center', yanchor='top'))
            col1.plotly_chart(fig5)

            # Visualization 8 - USE OF STATISTICS (Pie Chart)
            fig8 = px.pie(filtered_df, names='USE OF STATISTICS', title='Use of Statistics Distribution')
            fig8.update_layout(height=500, width=500)
            fig8.update_layout(margin=dict(l=20, r=20))
            fig8.update_layout(title=dict(x=0.5, y=1, xanchor='center', yanchor='top'))
            col2.plotly_chart(fig8)

            # Visualization 9 - CITE 'UN-habitat 2003 STUDY ON SLUMS' (Pie Chart)
            fig9 = px.pie(filtered_df, names='CITE \'UN-habitat 2003 STUDY ON SLUMS\'', title='Cite UN-habitat 2003 Study on Slums Distribution')
            fig9.update_layout(height=500, width=500)
            fig9.update_layout(margin=dict(l=20, r=20))
            fig9.update_layout(title=dict(x=0.5, y=1, xanchor='center', yanchor='top'))
            col1.plotly_chart(fig9)

            # Visualization 10 - CATEGORISATION OF PUBLICATION (Pie Chart)
            fig10 = px.pie(filtered_df, names='CATEGORISATION OF PUBLICATION( descriptive, diagnostic(identify problems), Prescriptive, theoretical)', title='Categorization of Publication Distribution')
            fig10.update_layout(height=500, width=500)
            fig10.update_layout(margin=dict(l=35, r=10))
            fig10.update_layout(title=dict(x=0.5, y=1, xanchor='center', yanchor='top'))
            col2.plotly_chart(fig10)

            #Visualization for Year of Publication
            col1.plotly_chart(fig3)

        else:
            st.warning("No data found for the selected filters.")
        

    elif page == "Word Cloud":
        st.markdown("<h2 style='text-align: center; padding-top: 1rem;'>Word Clouds</h2>", unsafe_allow_html=True)
        # Word Cloud
        st.markdown("""
        <div style="display: flex; justify-content: center; padding-top: 1rem;">
            <h3>Adjectival Phrases Word Cloud</h3>
        </div>
    """, unsafe_allow_html=True)
        text_data = 'DESCRIPTIVE TERMS/ADJECTIVES/ADJECTIVAL PHRASES (e.g. filthy, dirty, undesirable)'  
        wordcloud = WordCloud(width=600, height=400, background_color='white', mode='RGBA').generate(' '.join(filtered_df[text_data].dropna()))
        plt.figure(figsize=(10, 5))
        plt.imshow(wordcloud, interpolation='mitchell')
        plt.axis('off')
        st.set_option('deprecation.showPyplotGlobalUse', False)
        st.pyplot()


        col3, col4 = st.columns((2))

        col3.subheader('Descriptive studies Word Cloud')
        text_data = 'Descriptive studies( nature of slums(conceptualization), health concerns, housing, water, sanitation, nutrition,land, approaches to slum Upgrading)'  
        wordcloud = WordCloud(width=400, height=400, background_color='white', mode='RGBA').generate(' '.join(filtered_df[text_data].dropna()))
        plt.figure(figsize=(10, 5))
        plt.imshow(wordcloud, interpolation='bilinear')
        plt.axis('off')
        st.set_option('deprecation.showPyplotGlobalUse', False)
        col3.pyplot()

        col3.subheader('Prescriptive Studies Word Cloud')
        text_data = 'Prescriptive[inclusivity/tackling marginalisation, slum upgrading/improvement, security improvement,'  
        wordcloud = WordCloud(width=400, height=400, background_color='white', mode='RGBA').generate(' '.join(filtered_df[text_data].dropna()))
        plt.figure(figsize=(10, 5))
        plt.imshow(wordcloud, interpolation='bilinear')
        plt.axis('off')
        st.set_option('deprecation.showPyplotGlobalUse', False)
        col3.pyplot()

        col4.subheader('Diagnostic Studies Word Cloud')
        text_data = 'Diagnostic (strategies, policy gaps/challenges, tenure concerns, public health concerns, environmental concerns, marginalisation(inclusion, exclusion)]'  
        wordcloud = WordCloud(width=400, height=400, background_color='white', mode='RGBA').generate(' '.join(filtered_df[text_data].dropna()))
        plt.figure(figsize=(10, 5))
        plt.imshow(wordcloud, interpolation='bilinear')
        plt.axis('off')
        st.set_option('deprecation.showPyplotGlobalUse', False)
        col4.pyplot()

        col4.subheader('Theoretical Studies Word Cloud')
        text_data = 'Theoretical( theories, liberal/neo-liberal theories, colonial/post colonial theories, political economy)'  
        wordcloud = WordCloud(width=400, height=400, background_color='white', mode='RGBA').generate(' '.join(filtered_df[text_data].dropna()))
        plt.figure(figsize=(10, 5))
        plt.imshow(wordcloud, interpolation='bilinear')
        plt.axis('off')
        st.set_option('deprecation.showPyplotGlobalUse', False)
        col4.pyplot()

    else:
        st.warning('No Data Available')
            
else:
    st.warning("Please upload a CSV file.")


