import type {ReactNode} from 'react';
import clsx from 'clsx';
import Heading from '@theme/Heading';
import styles from './styles.module.css';

type FeatureItem = {
  title: string;
  Svg: React.ComponentType<React.ComponentProps<'svg'>>;
  description: ReactNode;
};

const FeatureList: FeatureItem[] = [
  {
    title: '🛡️ Secure & Reliable',
    Svg: require('@site/static/img/undraw_docusaurus_mountain.svg').default,
    description: (
      <>
        PS365 provides secure PowerShell functions to manage your Microsoft 365 tenant
        with confidence. Built following Microsoft best practices and security guidelines.
      </>
    ),
  },
  {
    title: '⚡ Powerful Automation',
    Svg: require('@site/static/img/undraw_docusaurus_tree.svg').default,
    description: (
      <>
        Automate complex Microsoft 365 administration tasks with ease.
        From Exchange Online to Azure AD, streamline your tenant management workflows.
      </>
    ),
  },
  {
    title: '📚 Well Documented',
    Svg: require('@site/static/img/undraw_docusaurus_react.svg').default,
    description: (
      <>
        Every function comes with detailed documentation, examples, and parameter descriptions.
        Get up and running quickly with comprehensive guides and usage scenarios.
      </>
    ),
  },
];

function Feature({title, Svg, description}: FeatureItem) {
  return (
    <div className={clsx('col col--4')}>
      <div className="text--center">
        <Svg className={styles.featureSvg} role="img" />
      </div>
      <div className="text--center padding-horiz--md">
        <Heading as="h3">{title}</Heading>
        <p>{description}</p>
      </div>
    </div>
  );
}

export default function HomepageFeatures(): ReactNode {
  return (
    <>
      <section className={styles.features}>
        <div className="container">
          <div className="row">
            {FeatureList.map((props, idx) => (
              <Feature key={idx} {...props} />
            ))}
          </div>
        </div>
      </section>
      
      <section className="margin-vert--lg">
        <div className="container">
          <div className="row">
            <div className="col col--8 col--offset-2">
              <div className="text--center margin-bottom--lg">
                <Heading as="h2">🚀 Quick Start</Heading>
                <p>Get started with PS365 in just a few steps</p>
              </div>
              <div className="card">
                <div className="card__header">
                  <h3>💻 Installation</h3>
                </div>
                <div className="card__body">
                  <pre><code>Install-Module -Name PS365 -Scope CurrentUser</code></pre>
                  <p>Install PS365 directly from the PowerShell Gallery</p>
                </div>
              </div>
            </div>
          </div>
        </div>
      </section>
    </>
  );
}
