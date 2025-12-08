const path = require('path');
const CopyWebpackPlugin = require('copy-webpack-plugin');

module.exports = {
  entry: {
    taskpane: './src/taskpane/taskpane.ts',
    commands: './src/commands/commands.ts'
  },
  output: {
    filename: '[name].js',
    path: path.resolve(__dirname, 'dist')
  },
  resolve: {
    extensions: ['.ts', '.tsx', '.js', '.json']
  },
  module: {
    rules: [
      {
        test: /\.tsx?$/,
        use: 'ts-loader',
        exclude: /node_modules/
      },
      {
        test: /\.html$/,
        use: 'html-loader'
      },
      {
        test: /\.css$/,
        use: ['style-loader', 'css-loader']
      }
    ]
  },
  plugins: [
    new CopyWebpackPlugin({
      patterns: [
        { from: 'src/taskpane/taskpane.html', to: 'taskpane.html' },
        { from: 'assets', to: 'assets' },
        { from: 'manifest.xml', to: 'manifest.xml' }
      ]
    })
  ],
  devServer: {
    static: path.join(__dirname, 'dist'),
    hot: true,
    port: 3000,
    devMiddleware: {
      writeToDisk: true
    }
  }
};
