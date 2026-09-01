declare module '*.module.scss.css' {
  const styles: { [key: string]: string };
  export default styles;
}

declare module '*.module.scss' {
  const styles: { [key: string]: string };
  export default styles;
}

declare module '*.scss' {
  const content: { [className: string]: string };
  export default content;
}
