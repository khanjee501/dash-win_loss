from turtle import numinput
import pandas as pd
import plotly.express as px
import dash
from dash import html
from dash import dcc
from dash import Input, Output
import dash_bootstrap_components as dbc
import statsmodels
import plotly.graph_objects as go
from plotly.subplots import make_subplots


df = pd.read_excel('DATA.xlsx',
                   sheet_name=['Win Loss', 'Bubble chart', 'Pie chart', 'Bar chart',
                               'Radar chart', 'Scatterplot 1', 'Scatterplot 2',
                               'Scatterplot 3', 'Scatterplot 4', 'Scatterplot 5'])

# print(df)

df_winloss = df.get('Win Loss')
df_bubble = pd.read_excel('Bubble chart (Updated.Xlsx')
df_bubble.fillna(0, inplace=True)
df_pie = df.get('Pie chart')
df_bar = df.get('Bar chart')
df_radar = df.get('Radar chart')
df_scatter_1 = df.get('Scatterplot 1')
df_scatter_2 = df.get('Scatterplot 2')
df_scatter_3 = df.get('Scatterplot 3')
df_scatter_4 = df.get('Scatterplot 4')
df_scatter_5 = df.get('Scatterplot 5')

df_pie = df_pie[df_pie['Year'] == 'TOTAL']
df_pie.reset_index(drop=True, inplace=True)
s_pie = df_pie.stack()
df_pie = pd.DataFrame(s_pie)
df_pie.columns = df_pie.iloc[0]
df_pie = df_pie.iloc[1:, :]
df_pie = df_pie.reset_index(level=[0, 1])
df_pie = df_pie.reset_index(level=[0])
df_pie = df_pie.rename(columns={'level_1': 'Condition'})

df_bar.fillna(0, inplace=True)
df_bar = df_bar.groupby('Year').agg(
    {'Win': 'sum', 'Loss': 'sum'}).reset_index()
df_bar['Loss'] = df_bar['Loss'].round(1)

df_radar = df_radar.transpose()
df_radar.columns = df_radar.iloc[0]
df_radar = df_radar.iloc[1:, :]

win = pd.read_excel('Win.xlsx')
loss = pd.read_excel('Loss.xlsx')

colors = {

    'bgcolor': '#1f2c56',

}


external_stylesheets = [dbc.themes.BOOTSTRAP]

app = dash.Dash(__name__, external_stylesheets=external_stylesheets)

app.layout = dbc.Container([

    dbc.Row([

        html.Div(id='none'),
        dbc.Col(
            html.Div([

                dcc.Graph(id='pie-chart',
                          style={'height': '50vh', 'margin-top': '20px'})

            ]), width=5
        ),

        dbc.Col(
            html.Div([
                dcc.Graph(id='bubble-chart',
                          style={'height': '50vh', 'margin-top': '20px'})
            ]), width=7
        ),

    ], style={"height": "50vh"}, className='g-2'),

    dbc.Row([
        dbc.Col(
            html.Div([
                dcc.Graph(id='radar-chart', figure={},
                          style={"height": "25vh", 'margin-top': '50px'})
            ]), width=4
        ),
        dbc.Col(
            html.Div([
                dcc.Graph(id='jr-radar-chart', figure={},
                          style={"height": "25vh", 'margin-top': '50px'})
            ]), width=4
        ),
        dbc.Col(
            html.Div([
                dcc.Graph(id='sr-radar-chart', figure={},
                          style={"height": "25vh", 'margin-top': '50px'})
            ]), width=4
        ),
    ], style={'height': '30vh'}, className='g-2'),


    dbc.Row([
        dbc.Col([

            html.Div([
                dcc.Graph(id='win-bar-chart',
                          style={'height': '50vh', 'margin-top': '50px'})
            ]),


        ], width=6),
        dbc.Col([

            html.Div([
                dcc.Graph(id='loss-bar-chart',
                          style={'height': '50vh', 'margin-top': '50px'})
            ]),

        ], width=6),

    ], style={'height': '50vh'}, className='g-2'),

    dbc.Row([
        dbc.Col([
            html.Div([
                dcc.Graph(id='scatter-one',
                          style={"height": "40vh", 'margin-top': '70px'})
            ]),
        ], width=6),
        dbc.Col(
            html.Div([
                dcc.Graph(
                    id='scatter-three', style={"height": "40vh", 'margin-top': '70px'})
            ]), width=6
        ),

    ], style={'height': '40vh'}, className='g-2'),

    dbc.Row([
        dbc.Col([
            html.Div([
                dcc.Graph(id='scatter-two',
                          style={"height": "40vh", 'margin-top': '90px'})
            ]),
        ], width=6),
        dbc.Col(
            html.Div([
                dcc.Graph(
                    id='scatter-four', style={"height": "40vh", 'margin-top': '90px'})
            ]), width=6
        ),

    ], style={'height': '40vh'}, className='g-2'),

    dbc.Row([
        dbc.Col(
            html.Div([
                dcc.Graph(id='main-scatterplot',
                          style={'height': '70vh', 'margin-top': '110px'})
            ]), width=12
        ),
    ], style={'height': '70vh'}, className='g-2')

], fluid=True,
    style={"height": "100vh"}
)


@app.callback(
    Output('pie-chart', 'figure'),
    Input('none', 'children')
)
def pie_chart(none):
    fig = px.pie(df_pie, names='Condition', values='TOTAL', color='Condition', hole=.4,
                 color_discrete_map={'Win': '#2E8B57',
                                     'Loss': '#CD5C5C', 'No bids': '#296D98'},
                 opacity=0.7)
    fig.update_layout(margin=dict(l=20, r=20, t=20, b=20), font_color='white',
                      plot_bgcolor=colors['bgcolor'], paper_bgcolor=colors['bgcolor'])
    return fig


@app.callback(
    Output('bubble-chart', 'figure'),
    Input('none', 'children')
)
def bubble_chart(none):

    try:
        fig = px.scatter(df_bubble, x='Year', y='Win', opacity=0.7, size='Win', size_max=40,
                         color_discrete_sequence=['#2E8B57'], template='plotly_white')
        # fig.update_traces(marker={'size': 50})
        trace1 = px.scatter(df_bubble, x='Year', y='Loss', opacity=0.7, size='Loss', size_max=40,
                            color_discrete_sequence=['#CD5C5C'])
        # trace1.update_traces(marker={'size': 25})
        trace2 = px.scatter(df_bubble, x='Year', y='No bids', opacity=0.7, size='No bids', size_max=40,
                            color_discrete_sequence=['#296D98'])
        # trace2.update_traces(marker={'size': 35})
        for scatter in trace1["data"]:
            fig.add_trace(scatter)
        for scatter in trace2["data"]:
            fig.add_trace(scatter)
        fig.update_layout(margin=dict(l=0, r=0, t=0, b=0), yaxis_title='Values',
                          xaxis_title='Date', font_color='white',
                          plot_bgcolor=colors['bgcolor'], paper_bgcolor=colors['bgcolor'],
                          xaxis={'showgrid': False, 'zeroline': False},
                          yaxis={'showgrid': False, 'zeroline': False})
        return fig
    except Exception:
        fig = px.scatter(df_bubble, x='Year', y='Win', opacity=0.7, size='Win',
                         color_discrete_sequence=['#2E8B57'], template='plotly_white')
        # fig.update_traces(marker={'size': 50})
        trace1 = px.scatter(df_bubble, x='Year', y='Loss', opacity=0.7, size='Loss',
                            color_discrete_sequence=['#CD5C5C'])
        # trace1.update_traces(marker={'size': 25})
        trace2 = px.scatter(df_bubble, x='Year', y='No bids', opacity=0.7, size='No bids',
                            color_discrete_sequence=['#296D98'])
        # trace2.update_traces(marker={'size': 35})
        for scatter in trace1["data"]:
            fig.add_trace(scatter)
        for scatter in trace2["data"]:
            fig.add_trace(scatter)
        fig.update_layout(margin=dict(l=0, r=0, t=0, b=0), yaxis_title='Values',
                          xaxis_title='Date', font_color='white',
                          plot_bgcolor=colors['bgcolor'], paper_bgcolor=colors['bgcolor'],
                          xaxis={'showgrid': False, 'zeroline': False},
                          yaxis={'showgrid': False, 'zeroline': False})
        return fig
        


@app.callback(
    Output('scatter-one', 'figure'),
    Input('none', 'children')
)
def scatter_one(none):

    try:
        fig = px.scatter(df_scatter_1, x='Year', y='Evaluate score', trendline='ols', log_y=True,
                         color_discrete_sequence=['cyan'], opacity=0.7)
        fig.update_layout(margin=dict(l=0, r=0, t=0, b=0), yaxis_title='Values',
                          xaxis_title='Date', autosize=False,
                          plot_bgcolor=colors['bgcolor'], paper_bgcolor=colors['bgcolor'],
                          xaxis={'showgrid': False, 'zeroline': False},
                          yaxis={'showgrid': False, 'zeroline': False},
                          font_color='white',)
        return fig
    except Exception:
        fig = px.scatter(df_scatter_1, x='Year', y='Evaluate score', trendline='ols', log_y=True,
                         color_discrete_sequence=['cyan'], opacity=0.7)
        fig.update_layout(margin=dict(l=0, r=0, t=0, b=0), yaxis_title='Values',
                          xaxis_title='Date', autosize=False,
                          plot_bgcolor=colors['bgcolor'], paper_bgcolor=colors['bgcolor'],
                          xaxis={'showgrid': False, 'zeroline': False},
                          yaxis={'showgrid': False, 'zeroline': False},
                          font_color='white',)
        return fig


@app.callback(
    Output('scatter-two', 'figure'),
    Input('none', 'children')
)
def scatter_two(none):

    try:
        fig = px.scatter(df_scatter_2, x='Year', trendline='ols', log_y=True,
                         y='Evaluate score', color_discrete_sequence=['cyan'], opacity=0.7)
        fig.update_layout(margin=dict(l=0, r=0, t=0, b=0), yaxis_title='Values',
                          xaxis_title='Date', autosize=False,
                          plot_bgcolor=colors['bgcolor'], paper_bgcolor=colors['bgcolor'],
                          xaxis={'showgrid': False, 'zeroline': False},
                          yaxis={'showgrid': False, 'zeroline': False},
                          font_color='white')
        return fig
    except Exception:
        fig = px.scatter(df_scatter_2, x='Year', trendline='ols', log_y=True,
                         y='Evaluate score', color_discrete_sequence=['cyan'], opacity=0.7)
        fig.update_layout(margin=dict(l=0, r=0, t=0, b=0), yaxis_title='Values',
                          xaxis_title='Date', autosize=False,
                          plot_bgcolor=colors['bgcolor'], paper_bgcolor=colors['bgcolor'],
                          xaxis={'showgrid': False, 'zeroline': False},
                          yaxis={'showgrid': False, 'zeroline': False},
                          font_color='white')
        return fig


@app.callback(
    Output('scatter-three', 'figure'),
    Input('none', 'children')
)
def scatter_three(none):

    try:
        fig = px.scatter(df_scatter_3, x='Year', y='Evaluate score', trendline='ols', log_y=True,
                         color_discrete_sequence=['cyan'], opacity=0.7)
        fig.update_layout(margin=dict(l=0, r=0, t=0, b=0), yaxis_title='Values',
                          xaxis_title='Date', autosize=False,
                          plot_bgcolor=colors['bgcolor'], paper_bgcolor=colors['bgcolor'],
                          xaxis={'showgrid': False, 'zeroline': False},
                          yaxis={'showgrid': False, 'zeroline': False},
                          font_color='white')
        return fig
    except Exception:
        fig = px.scatter(df_scatter_3, x='Year', y='Evaluate score', trendline='ols', log_y=True,
                         color_discrete_sequence=['cyan'], opacity=0.7)
        fig.update_layout(margin=dict(l=0, r=0, t=0, b=0), yaxis_title='Values',
                          xaxis_title='Date', autosize=False,
                          plot_bgcolor=colors['bgcolor'], paper_bgcolor=colors['bgcolor'],
                          xaxis={'showgrid': False, 'zeroline': False},
                          yaxis={'showgrid': False, 'zeroline': False},
                          font_color='white')
        return fig


@app.callback(
    Output('scatter-four', 'figure'),
    Input('none', 'children')
)
def scatter_four(none):

    try:
        fig = px.scatter(df_scatter_4, x='Year', y='Evaluate score', trendline='ols', log_y=True,
                         color_discrete_sequence=['cyan'], opacity=0.7)
        fig.update_layout(margin=dict(l=0, r=0, t=0, b=0), yaxis_title='Values',
                          xaxis_title='Date', autosize=False,
                          plot_bgcolor=colors['bgcolor'], paper_bgcolor=colors['bgcolor'],
                          xaxis={'showgrid': False, 'zeroline': False},
                          yaxis={'showgrid': False, 'zeroline': False},
                          font_color='white')
        return fig
    except Exception:
        fig = px.scatter(df_scatter_4, x='Year', y='Evaluate score', trendline='ols', log_y=True,
                         color_discrete_sequence=['cyan'], opacity=0.7)
        fig.update_layout(margin=dict(l=0, r=0, t=0, b=0), yaxis_title='Values',
                          xaxis_title='Date', autosize=False,
                          plot_bgcolor=colors['bgcolor'], paper_bgcolor=colors['bgcolor'],
                          xaxis={'showgrid': False, 'zeroline': False},
                          yaxis={'showgrid': False, 'zeroline': False},
                          font_color='white')
        return fig


@app.callback(
    Output('win-bar-chart', 'figure'),
    Input('none', 'children')
)
def bar_chart(none):

    fig = make_subplots(specs=[[{"secondary_y": True}]])

    fig.add_trace(
        go.Bar(x=win['Year'], y=win['Win'], marker_color='#2E8B57', opacity=0.7,
               ),
        secondary_y=False,
    )

    fig.add_trace(
        go.Scatter(x=win['Year'], y=win['Win'], marker_color='green', opacity=0.7, line=dict(width=4),
                   text=win['Win'], mode='lines+text', textposition='top center'),
        secondary_y=True

    )

    fig.update_layout(
        margin=dict(l=0, r=0, t=0, b=0), yaxis_title='Values', barmode='group',
        xaxis_title='Date',
        plot_bgcolor=colors['bgcolor'], paper_bgcolor=colors['bgcolor'],
        xaxis=dict(showgrid=False),
        yaxis=dict(showgrid=False),
        font_color='white',
        bargap=0.30, bargroupgap=0.0,
        showlegend=False
    )
    fig.update_yaxes(secondary_y=False, range=[0, 15])
    fig.update_yaxes(secondary_y=True, range=[0, 10])
    fig.update_traces(textfont_size=12, cliponaxis=False)
    fig['layout']['yaxis2']['showgrid'] = False

    return fig


@app.callback(
    Output('loss-bar-chart', 'figure'),
    Input('none', 'children')
)
def bar_chart(none):

    fig = make_subplots(specs=[[{"secondary_y": True}]])

    fig.add_trace(
        go.Bar(x=loss['Year'], y=loss['Loss'], marker_color='indianred', opacity=0.7,
               ),
        secondary_y=False,
    )

    fig.add_trace(
        go.Scatter(x=loss['Year'], y=loss['Loss'], marker_color='red', text=win['Win'],
                   opacity=0.7, line=dict(width=4), mode='lines+text', textposition='top center'),
        secondary_y=True

    )
    fig.update_yaxes(secondary_y=False, range=[0, 15])
    fig.update_yaxes(secondary_y=True, range=[0, 10])
    fig.update_traces(textfont_size=12, cliponaxis=False)
    fig['layout']['yaxis2']['showgrid'] = False

    fig.update_layout(
        margin=dict(l=0, r=0, t=0, b=0), yaxis_title='Values', barmode='group',
        xaxis_title='Date',
        plot_bgcolor=colors['bgcolor'], paper_bgcolor=colors['bgcolor'],
        xaxis={'showgrid': False, 'zeroline': False},
        yaxis={'showgrid': False, 'zeroline': False},
        font_color='white',
        bargap=0.30, bargroupgap=0.0,
        showlegend=False
    )

    return fig


@app.callback(
    Output('main-scatterplot', 'figure'),
    Input('none', 'children')
)
def bubble_chart(none):

    try:
        fig = px.scatter(df_scatter_5, x='Year', y='Evaluate score (KMD)', trendline='ols', log_y=True,
                         color_discrete_sequence=['indianred'], opacity=0.7, size_max=20, size='Evaluate score (KMD)')
        # fig.update_traces(marker={'size': 20})
        fig.add_traces(list(px.scatter(df_scatter_5, x='Year', y='Evaluate score (NET COMPANY)', trendline='ols', log_y=True,
                            color_discrete_sequence=['#296D98'], opacity=0.7, size_max=20, 
                            size='Evaluate score (NET COMPANY)').select_traces()))
        # trace1.update_traces(marker={'size': 20})
        fig.add_traces(list(px.scatter(df_scatter_5, x='Year', y='Evaluate score (NNIT)', trendline='ols', log_y=True,
                            color_discrete_sequence=['cyan'], opacity=0.7, size_max=20, 
                            size='Evaluate score (NNIT)').select_traces()))
        # trace2.update_traces(marker={'size': 20})
        # for scatter in trace1["data"]:
        #     fig.add_trace(scatter)
        # for scatter in trace2["data"]:
        #     fig.add_trace(scatter)
        fig.update_layout(margin=dict(l=0, r=0, t=0, b=0), yaxis_title='Values',
                          xaxis_title='Date', font_color='white',
                          plot_bgcolor=colors['bgcolor'], paper_bgcolor=colors['bgcolor'],
                          xaxis={'showgrid': False, 'zeroline': False},
                          yaxis={'showgrid': False, 'zeroline': False})
        return fig
    except Exception:
        fig = px.scatter(df_scatter_5, x='Year', y='Evaluate score (KMD)', trendline='ols', log_y=True,
                         color_discrete_sequence=['#964000'], opacity=0.7,)
        # fig.update_traces(marker={'size': 20})
        fig.add_traces(list(px.scatter(df_scatter_5, x='Year', y='Evaluate score (NET COMPANY)', trendline='ols', log_y=True,
                            color_discrete_sequence=['hotpink'], opacity=0.7,).select_traces()))
        # trace1.update_traces(marker={'size': 20})
        fig.add_traces(list(px.scatter(df_scatter_5, x='Year', y='Evaluate score (NNIT)', trendline='ols', log_y=True,
                            color_discrete_sequence=['orange'], opacity=0.7,).select_traces()))
        # trace2.update_traces(marker={'size': 20})
        # for scatter in trace1["data"]:
        #     fig.add_trace(scatter)
        # for scatter in trace2["data"]:
        #     fig.add_trace(scatter)
        fig.update_layout(margin=dict(l=0, r=0, t=0, b=0), yaxis_title='Values',
                          xaxis_title='Date', font_color='white',
                          plot_bgcolor=colors['bgcolor'], paper_bgcolor=colors['bgcolor'],
                          xaxis={'showgrid': False, 'zeroline': False},
                          yaxis={'showgrid': False, 'zeroline': False})
        return fig

        


@app.callback(
    Output('radar-chart', 'figure'),
    Input('none', 'children')
)
def radar_chart(none):

    try:
        fig = px.line_polar(df_radar, r=df_radar['Senior consultant'],
                            theta=df_radar.index, line_close=True)
        fig.update_traces(fill='toself')
        fig.update_polars(bgcolor=colors['bgcolor'])
        fig.update_layout(margin=dict(l=30, r=30, t=30, b=30), yaxis_title=None,
                        xaxis_title=None,
                        plot_bgcolor=colors['bgcolor'], paper_bgcolor=colors['bgcolor'],
                        xaxis={'showgrid': False, 'zeroline': False},
                        yaxis={'showgrid': False, 'zeroline': False},
                        font_color='white',)
        return fig
    except Exception:
        fig = px.line_polar(df_radar, r=df_radar['Senior consultant'],
                            theta=df_radar.index, line_close=True)
        fig.update_traces(fill='toself')
        fig.update_polars(bgcolor=colors['bgcolor'])
        fig.update_layout(margin=dict(l=30, r=30, t=30, b=30), yaxis_title=None,
                        xaxis_title=None,
                        plot_bgcolor=colors['bgcolor'], paper_bgcolor=colors['bgcolor'],
                        xaxis={'showgrid': False, 'zeroline': False},
                        yaxis={'showgrid': False, 'zeroline': False},
                        font_color='white',)



@app.callback(
    Output('sr-radar-chart', 'figure'),
    Input('none', 'children')
)
def sr_radar_chart(none):

    try:
        fig = px.line_polar(df_radar, r=df_radar['Foreign Senior consultant'],
                            theta=df_radar.index, line_close=True)
        fig.update_traces(fill='toself')
        fig.update_polars(bgcolor=colors['bgcolor'])
        fig.update_layout(margin=dict(l=0, r=0, t=0, b=0), yaxis_title=None,
                        xaxis_title=None,
                        plot_bgcolor=colors['bgcolor'], paper_bgcolor=colors['bgcolor'],
                        xaxis={'showgrid': False, 'zeroline': False},
                        yaxis={'showgrid': False, 'zeroline': False},
                        font_color='white',)
        return fig
    except Exception:
        fig = px.line_polar(df_radar, r=df_radar['Foreign Senior consultant'],
                            theta=df_radar.index, line_close=True)
        fig.update_traces(fill='toself')
        fig.update_polars(bgcolor=colors['bgcolor'])
        fig.update_layout(margin=dict(l=0, r=0, t=0, b=0), yaxis_title=None,
                        xaxis_title=None,
                        plot_bgcolor=colors['bgcolor'], paper_bgcolor=colors['bgcolor'],
                        xaxis={'showgrid': False, 'zeroline': False},
                        yaxis={'showgrid': False, 'zeroline': False},
                        font_color='white',)
        return fig


@app.callback(
    Output('jr-radar-chart', 'figure'),
    Input('none', 'children')
)
def jr_radar_chart(none):

    try:
        fig = px.line_polar(df_radar, r=df_radar['Junior consultant'],
                            theta=df_radar.index, line_close=True)
        fig.update_traces(fill='toself')
        fig.update_polars(bgcolor=colors['bgcolor'])
        fig.update_layout(margin=dict(l=30, r=30, t=30, b=30), yaxis_title=None,
                        xaxis_title=None,
                        plot_bgcolor=colors['bgcolor'], paper_bgcolor=colors['bgcolor'],
                        xaxis={'showgrid': False, 'zeroline': False},
                        yaxis={'showgrid': False, 'zeroline': False},
                        font_color='white',)
        return fig
    except Exception:
        fig = px.line_polar(df_radar, r=df_radar['Junior consultant'],
                            theta=df_radar.index, line_close=True)
        fig.update_traces(fill='toself')
        fig.update_polars(bgcolor=colors['bgcolor'])
        fig.update_layout(margin=dict(l=30, r=30, t=30, b=30), yaxis_title=None,
                        xaxis_title=None,
                        plot_bgcolor=colors['bgcolor'], paper_bgcolor=colors['bgcolor'],
                        xaxis={'showgrid': False, 'zeroline': False},
                        yaxis={'showgrid': False, 'zeroline': False},
                        font_color='white',)
        return fig


if __name__ == '__main__':
    app.run_server(debug=True)
